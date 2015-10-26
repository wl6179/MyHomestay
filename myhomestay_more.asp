<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
Select Case Request("action")
	Case "about_us"
	call sysshow()  '首先获得模版~...并替换$show_list$为注册条款信息；
	show=replace(show,"$show_list$","当前位置：<a href='index.asp'>首页</a>→关于我们<hr />"&oblog.setup(87,0))
	response.Write show&oblog.site_bottom
	Response.End '并且结束所有输出；
	
	Case "contact_us"
	call sysshow()  '首先获得模版~...并替换$show_list$为注册条款信息；
	show=replace(show,"$show_list$","当前位置：<a href='index.asp'>首页</a>→联系我们<hr />"&oblog.setup(88,0))
	response.Write show&oblog.site_bottom
	Response.End '并且结束所有输出；
	
	Case "cooperate_info"
	call sysshow()  '首先获得模版~...并替换$show_list$为注册条款信息；
	show=replace(show,"$show_list$","当前位置：<a href='index.asp'>首页</a>→合作代理<hr />"&oblog.setup(89,0))
	response.Write show&oblog.site_bottom
	Response.End '并且结束所有输出；
	
	Case "help"
	call sysshow()  '首先获得模版~...并替换$show_list$为注册条款信息；
	show=replace(show,"$show_list$","当前位置：<a href='index.asp'>首页</a>→帮助中心<hr />"&oblog.setup(90,0))
	response.Write show&oblog.site_bottom
	Response.End '并且结束所有输出；
	
	Case "especially"
	call sysshow()  '首先获得模版~...并替换$show_list$为注册条款信息；
	show=replace(show,"$show_list$","当前位置：<a href='index.asp'>首页</a>→特别提示<hr />"&oblog.setup(91,0))
	response.Write show&oblog.site_bottom
	Response.End '并且结束所有输出；
	
	Case "recommend"
	call sysshow()  '首先获得模版~...并替换$show_list$为注册条款信息；
	show=replace(show,"$show_list$","当前位置：<a href='index.asp'>首页</a>→推荐保护软件<hr />"&oblog.setup(92,0))
	response.Write show&oblog.site_bottom
	Response.End '并且结束所有输出；
	
	Case else 
	call sysshow()  '首先获得模版~...并替换$show_list$为注册条款信息；
	show=replace(show,"$show_list$","当前位置：<a href='index.asp'>首页</a>→欢迎光临我的住家网<hr />"&oblog.setup(86,0))
	response.Write show&oblog.site_bottom
	Response.End '并且结束所有输出；
End Select

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