<!--#include file="conn.asp"-->
<%
dim BrowserType,c_uname,c_upass,c_uurl
set BrowserType = Request.ServerVariables("HTTP_USER_AGENT")
c_uname=DecodeCookie(request.Cookies(cookies_name)("UserName"))
c_upass=DecodeCookie(request.Cookies(cookies_name)("Password"))
c_uurl=DecodeCookie(request.Cookies(cookies_name)("userurl"))
if left(c_uurl,1)<>"/" then
	c_uurl="http://"&c_uurl
end if
if instr(BrowserType,"Firefox") or instr(BrowserType,"Opera") or true_domain=1 then
%>
if (chkdiv('oblog_edit')) {
document.getElementById('oblog_edit').style.border='none';
document.getElementById('oblog_edit').style.background='none';
document.getElementById('oblog_edit').innerHTML='<textarea name="oblog_edittext"  rows="10" style="width:100%;"></textarea>';
}
function Verifycomment()
{
	document.all("edit").value=document.all("oblog_edittext").value;
	return true;
}
function reply_quote(id)
{
	document.all("oblog_edittext").value+="<div class='quote'><strong>以下引用"+document.all["n_"+id].innerHTML+"在"+document.all["t_"+id].innerHTML+"发布的评论:</strong><br /><br />"+document.all["c_"+id].innerHTML+"</div>\n"
	document.all("oblog_edittext").focus();
}
<%else%>
if (chkdiv('oblog_edit')) {
document.getElementById('oblog_edit').innerHTML='<ul> <select language="javascript" class="oblog_TBGen" id="FontSize" onchange="FormatText(\'fontsize\',this[this.selectedIndex].value);">\n	<option class="heading" selected>字号 \n	<option value="1">1 \n	<option value="2">2 \n	<option value="3">3 \n	<option value="4">4 \n	<option value="5">5 \n	<option value="6">6 \n	<option value="7">7</option>\n	</select> \n	<li class="oblog_Btn" TITLE="加粗" LANGUAGE="javascript" onclick="FormatText(\'bold\', \'\');ondrag=\'return false;\'" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\'; > \n	<img class="oblog_Ico" src="<%=blogurl%>images/bold.gif" /> </li>\n	<li class="oblog_Btn" TITLE="斜体" LANGUAGE="javascript" onclick="FormatText(\'italic\', \'\');ondrag=\'return false;\'" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\'; > \n	<img class="oblog_Ico" src="<%=blogurl%>images/italic.gif" /> </li>\n	<li class="oblog_Btn" TITLE="下划线" LANGUAGE="javascript" onclick="FormatText(\'underline\', \'\');ondrag=\'return false;\'" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\'; > \n	<img class="oblog_Ico" src="<%=blogurl%>images/underline.gif" /> </li>\n	<li class="oblog_Btn" TITLE="取消格式" LANGUAGE="javascript" onclick="FormatText(\'RemoveFormat\', \'\');ondrag=\'return false;\'" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\'; > \n	<img class="oblog_Ico" src="<%=blogurl%>images/removeformat.gif" /> </li>\n	<li class="oblog_Btn" TITLE="左对齐" NAME="Justify" LANGUAGE="javascript" onclick="FormatText(\'justifyleft\', \'\');ondrag=\'return false;\'" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\'; > \n	<img class="oblog_Ico" src="<%=blogurl%>images/aleft.gif" /> </li>\n	<li class="oblog_Btn" TITLE="居中" NAME="Justify" LANGUAGE="javascript" onclick="FormatText(\'justifycenter\', \'\');ondrag=\'return false;\'" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\'; > \n	<img class="oblog_Ico" src="<%=blogurl%>images/center.gif" /> </li>\n	<li class="oblog_Btn" TITLE="右对齐" NAME="Justify" LANGUAGE="javascript" onclick="FormatText(\'justifyright\', \'\');ondrag=\'return false;\'" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\'; > \n	<img class="oblog_Ico" src="<%=blogurl%>images/aright.gif" /> </li>\n	<li class="oblog_Btn" TITLE="插入表情" LANGUAGE="javascript" onclick="oblog_foremot()" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\'; > \n	<img class="oblog_Ico" src="<%=blogurl%>images/smiley.gif" /> </li>\n	<li id="forecolor" name=forecolor class="oblog_Btn" TITLE="字体颜色" LANGUAGE="javascript" onclick="oblog_foreColor();" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\'; > \n	<img class="oblog_Ico" src="<%=blogurl%>images/fgcolor.gif" /> </li>\n	<li id="backcolor" class="oblog_Btn" TITLE="字体背景颜色" LANGUAGE="javascript" onclick="oblog_backColor();ondrag=\'return false;\'" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\';> \n	<img class="oblog_Ico" src="<%=blogurl%>images/fbcolor.gif" /> </li>\n	<li class="oblog_Btn" TITLE="插入超级链接" LANGUAGE="javascript" onclick="oblog_forlink();ondrag=\'return false;\'" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\'; > \n	<img class="oblog_Ico" src="<%=blogurl%>images/wlink.gif" /> </li>\n	<li class="oblog_Btn" TITLE="去掉超级链接" LANGUAGE="javascript" onclick="FormatText(\'Unlink\');ondrag=\'return false;\'" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\'; > \n	<img class="oblog_Ico" src="<%=blogurl%>images/unlink.gif" /> </li>\n	<li class="oblog_Btn" TITLE="清理代码" LANGUAGE="javascript" onclick="oblog_CleanCode();ondrag=\'return false;\'" onmouseover=this.className=\'oblog_BtnMouseOverUp\'; onmouseout=this.className=\'oblog_Btn\';> \n	<img class="oblog_Ico" src="<%=blogurl%>images/cleancode.gif" /></li>\n	</ul>\n	<ul style="height:100%" id="PostiFrame"> <iframe class="oblog_Composition" ID="oblog_Composition" MARGINHEIGHT="5" MARGINWIDTH="5" width="100%" height="100%"></iframe> </ul>\n    <ul style="text-align:right;"><a href="javascript:oblog_Size(-360)"><img src="<%=blogurl%>images/minus.gif" border=\'0\' height="20" /></a> <a href="javascript:oblog_Size(330)"><img src="<%=blogurl%>images/plus.gif"  border=\'0\' height="20" /></a>\n	<li style="width:10px;"></li>\n	</ul>\n';
document.write('<script src="<%=blogurl%>images/DhtmlEdit.js"></script>');
document.write('<script src="<%=blogurl%>images/editor_s.asp?blogurl=<%=blogurl%>"></script>');
}
<%end if

Function DecodeCookie(Str)
        Dim i
        Dim StrArr, StrRtn
        StrArr = Split(Str, "a")
        For i = 0 To UBound(StrArr)
            If IsNumeric(StrArr(i)) = True Then
                StrRtn = ChrW(StrArr(i)) & StrRtn
            Else
                StrRtn = Str
                Exit Function
            End If
        Next
        DecodeCookie = StrRtn
End Function
%>
document.getElementById('UserName').value='<%=c_uname%>';
document.getElementById('Password').value='<%=c_upass%>';
document.getElementById('homepage').value='<%=c_uurl%>';