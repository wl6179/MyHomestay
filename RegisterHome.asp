<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
If Request("action")="protocol" Then
	call sysshow()  '首先获得模版~...并替换$show_list$为注册条款信息；
	show=replace(show,"$show_list$","当前位置：<a href='index.asp'>首页</a>→接待家庭注册条款<hr />"&oblog.setup(47,0))
	response.Write show&oblog.site_bottom
	Response.End '并且结束所有输出；	
End If
dim action,show_reg,chkregtime
chkregtime=30 '重复注册的时间间隔,单位：秒
if is_ot_user then
	if not IsObject(conn) then link_database
	response.Redirect(ot_regurl)
	response.End()
end if
action=trim(request.Form("action"))
call sysshow_3()  '从数据库首先获得03模版~...在往下，就是判断是处理注册 或者 显示出注册表单了；
select case action
	case "chkreg"
	call sub_chkreg()'处理注册;
	case "sub_showreg"
	call sub_showreg()';
	case else 
	call sub_showregSelect()'显示出注册选择类型表单了;
end select
show=replace(show,"$show_list$",show_reg)&oblog.site_bottom'将标签替$show_list$换为 该有的内容 show_reg 后 + 底部版权信息代码；
response.Write show'全面输出。

sub sub_showregSelect()

	show_reg=show_reg&"当前位置：<a href='index.asp'>首页</a> → 请您选择接待家庭种类<br /><br /><font color=#038ad7>请您选择接待家庭种类</font><br /><br />"
	show_reg=show_reg&"<form name='FamilyForm' method='post' action='RegisterHome.asp' onSubmit='return Check();'>"

	show_reg=show_reg&"<script language='JavaScript'>"
		show_reg=show_reg&"function buttontrue(){"
		show_reg=show_reg&"document.FamilyForm.SubmitInfo.disabled=false;"
		show_reg=show_reg&"}	"	
	show_reg=show_reg&"</script>"
	
	show_reg=show_reg&"<script language='JavaScript'>"
		show_reg=show_reg&"function Check(){"
		show_reg=show_reg&"if(document.FamilyForm.FamilyType.value==''){"
		show_reg=show_reg&"alert('请选择您申请的接待家庭类型');"
		show_reg=show_reg&"FamilyForm.FamilyType.focus();"
		show_reg=show_reg&"return false;"
		show_reg=show_reg&"}"
		show_reg=show_reg&"}"
	show_reg=show_reg&"</script>"

	
	  show_reg=show_reg&"<center><input name='FamilyType' type='checkbox' value='1' onClick='buttontrue();'>"
	  show_reg=show_reg&"申请免费接待家庭 "
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='2' onClick='buttontrue();'>"
	  show_reg=show_reg&"申请收费接待家庭 "
	  show_reg=show_reg&"<input name='FamilyType' type='checkbox' value='3' onClick='buttontrue();'>"
	  show_reg=show_reg&"<font color='#038ad7'>申请2008奥运接待家庭</font>"
	  
	  show_reg=show_reg&"<input name='action' type='hidden' value='sub_showreg' />"
	  show_reg=show_reg&"<br /><br /><input type='submit' disabled id='SubmitInfo' name='SubmitInfo' value=' 下一步(Next) &gt;&gt;' style='height:28px;width:150px;' ></center>"
	  
	  
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
	FamilyTypeRequest= "免费接待家庭"
	Case "2"
	FamilyTypeRequest= "收费接待家庭"
	Case "3"
	FamilyTypeRequest= "2008奥运接待家庭"
	Case "1, 2"
	FamilyTypeRequest= "免费/收费接待家庭"
	Case "1, 3"
	FamilyTypeRequest= "免费接待家庭/2008奥运接待家庭"
	Case "1, 2, 3"
	FamilyTypeRequest= "免费/收费/2008奥运接待家庭"
	Case "2, 3"
	FamilyTypeRequest= "收费接待家庭/2008奥运接待家庭"
    End Select
	'response.Write("("&FamilyTypeRequest&")") 

	show_reg=show_reg&"当前位置：<a href='index.asp'>首页</a> → 申请(免费/收费/2008奥运)接待家庭<br /><br /><center><h3 style='height:38px;color:#00A1E4;background:url(/images/HomestayHoltel_internationalization.gif) no-repeat 30px -1px; height:130px; padding-top:0px; line-height:88px; padding-left:200px; font-size:15px;'>&nbsp;申请接待家庭</h3><font color=#038ad7>您当前申请的是 "& FamilyTypeRequest &"</font></center>"

	if oblog.setup(15,0)=0 and session("adminname")="" then
		show_reg=show_reg&"当前系统已关闭注册。"
	else
	str_usertype="<select name=usertype id=usertype>"
	str_usertype=str_usertype&oblog.show_class("user",0,0)
    str_usertype=str_usertype&"</select><font color=#038ad7> *</font>"
	show_reg=show_reg&"<form name='oblogform' id='oblogform' method='post' action='RegisterHome.asp' onSubmit='return VerifySubmit()'>"& vbcrlf
	show_reg=show_reg&"<table width=90% border=0 cellspacing=0 cellpadding=0 align=center>"& vbcrlf
    
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=center style='color:#999999'>家庭帐号信息</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>家庭登录帐号：</div></td>"& vbcrlf
    show_reg=show_reg&"<td width=63% colspan=2> <input name=username type=text id=username size=15 maxlength=30><font color=#038ad7 > *</font>&nbsp;(如:homestay2008)&nbsp;<a href=""javascript:checkssn('checkssn.asp')"";>→查看此帐号是否被占用</a>"& vbcrlf 
    show_reg=show_reg&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		show_reg=show_reg&"<tr> "& vbcrlf
		show_reg=show_reg&"<td width=20% height=25><div align=left >域名：</div></td>"&vbcrlf
		show_reg=show_reg&"<td width=10% > <input name=domain type=text id=domain size=15 maxlength=30></td><td align=left >"&" <select name='user_domainroot'>"&oblog.type_domainroot("")&"</select><font color=#ff0000 > *</font>"& vbcrlf 
		show_reg=show_reg&"</td>"& vbcrlf
		show_reg=show_reg&"</tr>"& vbcrlf
	else
		show_reg=show_reg&"<input name=domain type=hidden id=domain size=15 maxlength=30><input type=hidden name='user_domainroot'>"& vbcrlf 
	end if
'	show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
'    show_reg=show_reg&"<td width=63% colspan=2 ><a href=""javascript:checkssn('checkssn.asp')"";>查看以上帐号是否被占用</a>"& vbcrlf 
'    show_reg=show_reg&"</td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>输入密码：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input style='width:9em;' name=password type=password id=password size=15 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>登录密码确认：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input style='width:9em;' name=repassword type=password id=repassword size=15 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
'    show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25><div align=left>密码提示问题：</div></td>"& vbcrlf
'    show_reg=show_reg&"<td colspan=2><input name=question type=text id=question size=30 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
'    show_reg=show_reg&"<tr> "& vbcrlf
'    show_reg=show_reg&"<td width=20% height=25><div align=left>密码提示答案：</div></td>"& vbcrlf
'    show_reg=show_reg&"<td colspan=2><input name=an type=text id=an size=30 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
'    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=center style='color:#999999'>家庭基本信息</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf

	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>所在地区(省/市)：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"&oblog.type_city("","")&"<font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>Email地址：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=email type=text size=35 maxlength=30><font color=#038ad7> *</font></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>真实姓名：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name=blogname type=text size=15 maxlength=30><font color=#038ad7> *</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;您的称呼：<select name=sex><option value=3>请选择</option><option value=0>女士</option><option value=1>先生</option></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>请选择学习范围：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"&str_usertype&"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>出生日期：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><select name='y' size=1>"

	show_reg=show_reg& oblog.y()&"</select>年<select name='m' size=1>"& oblog.m()&"</select>月<select name='d' size=1>"& oblog.d()&"</select>日</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>您的职业：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>"& oblog.type_job2("") &"</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>联系方式：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>电话：<input name=tel value='' size='15' maxlength='50' />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;手机：<input name=mobile  value='' size='15' maxlength='50' /></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>地址：<input name='address' value='' size='30' maxlength='250' />&nbsp;&nbsp;&nbsp;邮编：<input name=postnum value='' size='8' maxlength='6' /></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>家庭成员：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>家庭人数：<input id='familynumber_0' type='radio' name='familynumber' value='1' /><label for='familynumber_0'>1</label>&nbsp;&nbsp;<input id='familynumber_1' type='radio' name='familynumber' value='2' /><label for='familynumber_1'>2</label>&nbsp;&nbsp;<input id='familynumber_2' type='radio' name='familynumber' value='3' checked='checked' /><label for='familynumber_2'>3</label>&nbsp;&nbsp;<input id='familynumber_3' type='radio' name='familynumber' value='4' /><label for='familynumber_3'>4</label>&nbsp;&nbsp;<input id='familynumber_4' type='radio' name='familynumber' value='更多' /><label for='familynumber_4'>更多</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>家庭结构：<input name='familymember' type='checkbox' id='Father' value='父亲' title='请选择您的家庭结构，可多选'><label for='Father'>父亲</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Mother' value='母亲' title='请选择您的家庭结构，可多选'><label for='Mother'>母亲</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Husband' value='丈夫' title='请选择您的家庭结构，可多选'><label for='Husband'>丈夫</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Wife' value='妻子' title='请选择您的家庭结构，可多选'><label for='Wife'>妻子</label><br>&nbsp;&nbsp;<font color=#ffffff>Homestay</font><input name='familymember' type='checkbox' id='Son' value='儿子' title='请选择您的家庭结构，可多选'><label for='Son'>儿子</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Daughter' value='女儿' title='请选择您的家庭结构，可多选'><label for='Daughter'>女儿</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Brother' value='兄弟' title='请选择您的家庭结构，可多选'><label for='Brother'>兄弟</label>&nbsp;&nbsp;<input name='familymember' type='checkbox' id='Sister' value='姐妹' title='请选择您的家庭结构，可多选'><label for='Sister'>姐妹</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
                           
    
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>房屋介绍：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>结构：<select name='houseframe' id='houseframe1' title='请选择您的房屋结构'>"&_
                            "<option value='1室' selected>1室</option>"&_
                            "<option value='2室'>2室</option>"&_
                            "<option value='3室'>3室</option>"&_
                            "<option value='4室'>4室</option>"&_
                            "<option value='别墅'>别墅</option>"&_
                            "<option value='其他'>其他</option>"&_
                          "</select>&nbsp;<select name='houseframe' id='houseframe2'  title='请选择您的房屋结构'>"&_
                            "<option value='1厅' selected>1厅</option>"&_
                            "<option value='2厅'>2厅</option>"&_
                            "<option value='3厅'>3厅</option>"&_
                            "<option value='4厅'>4厅</option>"&_
                            "<option value='5厅'>5厅</option>"&_
                            "<option value='6厅'>6厅</option>"&_
                          "</select>&nbsp;<select name='houseframe' id='houseframe3' title='请选择您的房屋结构'>"&_
                            "<option value='1卫' selected>1卫</option>"&_
                            "<option value='2卫'>2卫</option>"&_
                            "<option value='3卫'>3卫</option>"&_
                            "<option value='4卫'>4卫</option>"&_
                            "<option value='5卫'>5卫</option>"&_
                            "<option value='6卫'>6卫</option>"&_
                          "</select>"&_
						  "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;面积："&_ 
                          "<select name='housespace' id='housespace' style='width:100px' title='请选择您的房屋面积'>"&_
                            "<option selected>请选择</option>"&_
                            "<option value='50-100'>50到100平方米</option>"&_
                            "<option value='100-150'>100到150平方米</option>"&_
                            "<option value='150-200'>150到200平方米</option>"&_
                            "<option value='above 200'>200平方米以上</option>"&_
                          "</select>"&_
						  "</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf 
	

	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>是否有宠物：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT type=radio value='有' name='pet' id='Pet1' onClick='ShowPet(1);' title='请选择您的宠物种类'><label for='Pet1'>有</label>&nbsp;<INPUT name=pet type=radio value='没有' checked id='Pet2' onClick='ShowPet(0);'><label for='Pet2'>没有</label>&nbsp;&nbsp;&nbsp;<span width='21%' id='PetTypes'></span></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf                      
    
                           
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25>对外教的特殊要求：</td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>外教性别：<input id='asksex_0' type='radio' name='asksex' value='男' /><label for='asksex_0'>男</label>&nbsp;&nbsp;&nbsp;<input id='asksex_1' type='radio' name='asksex' value='女' /><label for='asksex_1'>女</label>&nbsp;&nbsp;&nbsp;<input id='asksex_2' type='radio' name='asksex' value='无所谓' checked='checked' /><label for='asksex_2'>无所谓</label> &nbsp;&nbsp;&nbsp;</td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf                
                           
    show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2>是否会说中文：<input id='issaychinese_0' type='radio' name='issaychinese' value='是' /><label for='issaychinese_0'>是</label>&nbsp;&nbsp;&nbsp;<input id='issaychinese_1' type='radio' name='issaychinese' value='否' /><label for='issaychinese_1'>否</label>&nbsp;&nbsp;&nbsp;<input id='issaychinese_2' type='radio' name='issaychinese' value='无所谓' checked='checked' /><label for='issaychinese_2'>无所谓</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf 
	
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=center style='color:#999999'>家庭附加信息</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><hr></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	                    
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>如何知道我们：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input id='howknowsite_0' type='checkbox' name='howknowsite' value='报纸杂志' /><label for='howknowsite_0'>报纸杂志</label>&nbsp;&nbsp;<input id='howknowsite_1' type='checkbox' name='howknowsite' value='朋友介绍' /><label for='howknowsite_1'>朋友介绍</label>&nbsp;&nbsp;<input id='howknowsite_2' type='checkbox' name='howknowsite' value='网站' /><label for='howknowsite_2'>网站</label>&nbsp;&nbsp;<input id='howknowsite_3' type='checkbox' name='howknowsite' value='论坛' /><label for='howknowsite_3'>论坛</label>&nbsp;&nbsp;<input id='howknowsite_4' type='checkbox' name='howknowsite' value='搜索引擎' /><label for='howknowsite_4'>搜索引擎</label>&nbsp;&nbsp;<input id='howknowsite_5' type='checkbox' name='howknowsite' value='其它' /><label for='howknowsite_5'>其它</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>英语水平：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><input name='englishlevel' id='englishlevel_0' type='radio' value='不会' title='请选择您英语水平'><label for='englishlevel_0'>不会</label>&nbsp;&nbsp;<input name='englishlevel' id='englishlevel_1' type='radio' value='一般' checked title='请选择您英语水平'><label for='englishlevel_1'>一般</label>&nbsp;&nbsp;<input type='radio' name='englishlevel' id='englishlevel_2' value='中等' title='请选择您英语水平'><label for='englishlevel_2'>中等</label>&nbsp;&nbsp;<input type='radio' name='englishlevel' id='englishlevel_3' value='很好' title='请选择您英语水平'><label for='englishlevel_3'>很好</label>&nbsp;&nbsp;&nbsp;&nbsp;<strong>是否有私家车</strong>：<INPUT name=car id=car_0 type=radio value=有 title='请选择是否有私家车'><label for='car_0'>有</label>&nbsp;<INPUT name=car id=car_1 type=radio value=没有 checked title='请选择是否有私家车'><label for='car_1'>没有</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>卫生间的类型：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT name=toilet id=toilet_0 type=radio value=共用 checked title='请选择是否有独立的卫生间'><label for='toilet_0'>共用</label><INPUT type=radio value=房间独立 name=toilet id=toilet_1 title='请选择是否有独立的卫生间'><label for='toilet_1'>房间独立</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>有否电脑：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><INPUT type=radio value=有 name=computer id=computer_0 title='请选择是否有电脑'><label for='computer_0'>有</label>&nbsp;&nbsp;<INPUT name=computer id=computer_1 type=radio value=没有 checked title='请选择是否有电脑'><label for='computer_1'>没有</label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>有否宽带上网：</strong><INPUT type=radio value=有 name=internet id=internet_0 title='请选择是否有宽带可上网'><label for='internet_0'>有</label>&nbsp;&nbsp;<INPUT name=internet id=internet_1 type=radio value=没有 checked title='请选择是否有宽带可上网'><label for='internet_1'>没有</label></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left>客房描述(或备注)：</div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2><textarea name=remarkinfo cols=50 rows=3 id='remarkinfo' title='请输入您家客房介绍，是单人床还是双人床，有否衣柜，空调，写字台等'></textarea></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	
	
	
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf
	show_reg=show_reg&"<tr> "& vbcrlf
    show_reg=show_reg&"<td width=20% height=25><div align=left></div></td>"& vbcrlf
    show_reg=show_reg&"<td colspan=2> <input class='input_radio' name=passregtext id=passregtext type=radio value='1' checked><label for='passregtext'>同意</label>　<input class='input_radio' type=radio name=passregtext id=nopassregtext value='0'>不同意　&nbsp;<a href='RegisterHome.asp?action=protocol' target='_blank'> → 查看注册条款</a></td>"& vbcrlf
    show_reg=show_reg&"</tr>"& vbcrlf	
	
	if oblog.setup(57,0)=1 then
		show_reg=show_reg&"<tr> "& vbcrlf
    	show_reg=show_reg&"<td width=20% height=25><div align=left>验证码：</div></td>"& vbcrlf
   	 	show_reg=show_reg&"<td colspan=2><input name=codestr type=text size=6 maxlength=4> "&oblog.getcode&"</td>"& vbcrlf
   		show_reg=show_reg&"</tr>"& vbcrlf
	end if	
	show_reg=show_reg&"<tr>"& vbcrlf 
    show_reg=show_reg&"<td height=25 ></td><td colspan=2> <div align=left>"& vbcrlf
    show_reg=show_reg&"<br><input type=submit name=Submit value=' 提交申请接待家庭 ' style='height:21px;font-size:13px;font-weight:bold'>　"& vbcrlf
	show_reg=show_reg&"<input type='button' name='historybackwl' value='返回上一步' onclick='javascript:history.go(-1);' class=btxx style='height:21px;font-size:13px;font-weight:bold;cursor:hand;'><input type=hidden name=action value='chkreg'><input type=hidden name=FamilyType value='"& Request.Form("FamilyType") &"'><input type=hidden name=FamilyTypeRequest value='"& FamilyTypeRequest &"'>"& vbcrlf
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
		oblog.adderrstr("系统不允许从外部提交！")
		oblog.showerr
		exit sub
	end if
	chk_regtime()
	if oblog.setup(57,0)=1 then
		if not oblog.codepass then oblog.adderrstr("验证码错误，请刷新后重新输入！"):oblog.showerr
	end if
	dim rsreg
	dim regusername,regpassword,sex,question,answer,email,reguserlevel,userispass,blogname,usertype,nickname,re_regpassword,user_domain,user_domainroot
	Dim birthday,job,tel,address
	Dim mobile,postnum,familynumber,familymember,houseframe,housespace,pet,PetType,asksex,issaychinese,howknowsite,englishlevel,car,fitment,toilet,computer,internet,remarkinfo
	
	
	
	if oblog.chkiplock() then
		oblog.adderrstr("对不起！你的IP已被锁定,不允许注册！")
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
	question="您尚未设置密码提示，如无法登录请与管理员联系，并为您找回密码：）"
	answer="我家李慧颖"
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
	
	if regusername="" or oblog.strLength(regusername)>14 or oblog.strLength(regusername)<4 then oblog.adderrstr("用户名不能为空，并且不能大于14小于4！")
	if oblog.chk_regname(regusername) then oblog.adderrstr("用户名系统不允许注册！")
	if oblog.chk_badword(regusername)>0 then oblog.adderrstr("用户名中含有系统不允许的字符！")
	if en_nameisnum=0 and IsNumeric(regusername) then oblog.adderrstr("用户名不允许全部为数字！")
	if oblog.chkdomain(regusername)=false then oblog.adderrstr("用户名不合规范，只能使用小写字母，数字及下划线！")
	if trim(oblog.setup(4,0))<>"" and oblog.setup(12,0)=1 then
		if user_domain="" or oblog.strLength(user_domain)>14  then oblog.adderrstr("域名不能为空(不能大于14个字符)！")
		if user_domain<>request("old_userdomain") and oblog.strLength(user_domain)<4 then oblog.adderrstr("域名不能小于4个字符！")
		if oblog.chk_regname(user_domain) then oblog.adderrstr("此域名系统不允许注册！")
		if oblog.chk_badword(user_domain)>0 then oblog.adderrstr("域名中含有系统不允许的字符！")
		if oblog.chkdomain(user_domain)=false then oblog.adderrstr("域名不合规范，只能使用小写字母，数字！")
		if user_domainroot="" then oblog.adderrstr("域名根不能为空！")
	end if
	if regpassword="" or oblog.strLength(regpassword)>14 or oblog.strLength(regpassword)<4 then oblog.adderrstr("密码不能为空(不能大于14小于4)！")
	if re_regpassword="" then oblog.adderrstr("重复密码不能为空！")
	if regpassword<>re_regpassword then oblog.adderrstr("两次输入密码不同！")
	if question="" or oblog.strLength(question)>150 then oblog.adderrstr("找回密码提示问题不能为空，也不能大于150！")
	if answer="" or oblog.strLength(answer)>150 then oblog.adderrstr("找回密码问题答案不能为空，也不能大于150！")
	'if oblog.chk_regname(nickname) then oblog.adderrstr("此昵称系统不允许注册！")
	'if oblog.chk_badword(nickname)>0 then oblog.adderrstr("昵称中含有系统不允许的字符！")
	'if oblog.strLength(nickname)>50 then oblog.adderrstr("昵称不能不能大于50字符！")
	if blogname="" or oblog.strLength(blogname)>50 then oblog.adderrstr("真实姓名不能为空(不能大于50字符)！")
	if oblog.chk_badword(blogname)>0 then oblog.adderrstr("真实姓名中含有系统不允许的字符！")	
	if Instr(regusername,"=")>0 or Instr(regusername,"%")>0 or Instr(regusername,chr(32))>0 or Instr(regusername,"?")>0 or Instr(regusername,"&")>0 or Instr(regusername,";")>0 or Instr(regusername,",")>0 or Instr(regusername,"'")>0 or Instr(regusername,",")>0 or Instr(regusername,chr(34))>0 or Instr(regusername,chr(9))>0 or Instr(regusername,"")>0 or Instr(regusername,"$")>0 then oblog.adderrstr("用户名中含有非法字符！")
	'if oblog.setup(25,0)=0 and nickname<>"" then
	'	set rsreg=oblog.execute("select userid from oblog_user where nickname='"&oblog.filt_badstr(nickname)&"'")
	'	if not rsreg.eof or not rsreg.bof then oblog.adderrstr("系统中已经有这个昵称存在，请更改昵称！")		
	'end if
	if user_domain<>"" then
		set rsreg=oblog.execute("select userid from oblog_user where user_domain='"&oblog.filt_badstr(user_domain)&"' and user_domainroot='"&oblog.filt_badstr(user_domainroot)&"'")
		if not rsreg.eof or not rsreg.bof then oblog.adderrstr("系统中已经有这个域名存在，请更改域名！")
	end if
	
	
	If IsDate(birthday)=False Then oblog.adderrstr("您的生日格式尚未填好，请重新选择！")	

	
	
	
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
'wl20070816,发送注册成功短消息----------------------------------	Begin	
		dim sqlt,rst
		sqlt="select * from oblog_pm"
		set rst=server.createobject("adodb.recordset")
		rst.open sqlt,conn,1,3
		rst.addnew
		rst("incept")=oblog.Interceptstr(regusername,50)'wl---regusername用户名;
		rst("topic")=oblog.Interceptstr(oblog.filt_badword("恭喜您注册成功...欢迎来到我的住家网！"),100)
		rst("content")=oblog.setup(86,0)
		rst("sender")="我的住家网"
		rst.update
		rst.close
		set rst= nothing
'wl20070816,发送注册成功短消息----------------------------------	End		
		show_reg="当前位置：<a href='index.asp'>首页</a>→完成注册<hr />"
		oblog.CreateUserDir regusername,1
		if oblog.setup(16,0)=1 then
			show_reg=show_reg&"<ul><li><strong>恭喜！您已经成功注册为 </strong><font color=#038ad7>"& FamilyTypeRequest &"</font> ！</li></ul>"
			show_reg=show_reg&"10秒后自动转到系统首页。"
			show_reg=show_reg&"<script language=JavaScript>"
			show_reg=show_reg&"setTimeout(""window.location='index.asp'"",10000);"
			show_reg=show_reg&"</script>"
		else
			oblog.savecookie regusername,regpassword,0,user_domain&"."&user_domainroot
			show_reg=show_reg&"<ul><li><strong>恭喜！您已经注册成功！</strong></li>"
			show_reg=show_reg&"<li><a href='index.asp'>返回首页</a></li>"
			show_reg=show_reg&"<li><a href='HomestayBackctrl-index.asp?url=user_template.asp?u=new'>点击进入管理后台选择您的喜欢的页面风格(5秒后自动进入管理后台)</a></li></ul>"
			show_reg=show_reg&"<script language=JavaScript>"
			show_reg=show_reg&"setTimeout(""window.location='HomestayBackctrl-index.asp?url=user_template.asp?u=new'"",5000);"
			show_reg=show_reg&"</script>"
		end if
	else
		oblog.adderrstr("系统中已经有这个用户名存在，请更改用户名！")
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
			Response.Write("<script language=javascript>alert('"&chkregtime&"秒后才能重复注册。');window.history.back(-1);</script>")
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
        alert("您不同意注册条款，不能注册!");
	return false;
   }	 
   
 
		 
	uid = del_space(document.oblogform.username.value);
     if (uid.length == 0)
     {
        alert("请输入用户名!");
	return false;
     }
	 
	pwd = del_space(document.oblogform.password.value);
     if (pwd.length == 0)
     {
        alert("请输入密码!");
	return false;
     }
	 
	 pwd = del_space(document.oblogform.password.value);
     if (pwd.length < 6)
     {
        alert("密码要大于6个字符！");
	return false;
     }

	pwd2 = del_space(document.oblogform.repassword.value);
     if (pwd2!=pwd)
     {
        alert("两次输入的密码不一致！");
		return false;
     }
	 
//	 tishi = del_space(document.oblogform.question.value);
//     if (tishi.length == 0)
//     {
//        alert("请输入密码提示问题");
//        return false;
//     }
//	 
//	 tsda = del_space(document.oblogform.an.value);
//     if (tsda.length == 0)
//     {
//        alert("请输入密码提示问题答案！");
//        return false;
//     }
////	WL严重隔离！ city = del_space(document.all("city").value);
////     if (city.length == 0)
////     {
////        alert("请选择所在城市!");
////	return false;
////     }		
	 
	 
	email = del_space(document.all("email").value);
     if (email.length == 0)
     {
        alert("请输入Email!");
	return false;
     }
	 
     if (email.indexOf("@")==-1)
     {
        alert("Email地址无效!");
	return false;
     }
	 
	email = del_space(document.all("email").value);
     if (email.indexOf(".")==-1)
     {
        alert("Email地址无效!");
	return false;
     }
	 blogname = del_space(document.oblogform.blogname.value);
     if (blogname.length == 0)
     {
        alert("请输入您的真实姓名!");
	return false;
     }

	 
	if (document.oblogform.usertype.value == 0)
     {
        alert("请选择您的学习范围!");
	return false;
     }
	 
	 return true;
}
//-->
</SCRIPT>


<script>
	function ShowPet(ID){
		if(ID=="1"){
		window.PetTypes.innerHTML="<select id=\"PetType\" name=\"PetType\"><option value=\"\" selected>请选择<\/option><option value=\"狗\" >狗<\/option><option value=\"猫\">猫<\/option><option value=\"鸟\">鸟<\/option><option value=\"其他\">其他<\/option><\/select>";
		//document.oblogform.CheckPet.value="1"
		}
		if(ID=="0"){
		window.PetTypes.innerHTML="";
		//document.oblogform.CheckPet.value="0"
		}
		}
</script>