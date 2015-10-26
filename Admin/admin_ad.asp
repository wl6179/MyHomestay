<!--#include file="inc/inc_sys.asp"-->
<%
dim action,rs,site_placard
Action=trim(request("Action"))
if action="saveconfig" then
	dim fso,dirstr,ad_usertop,ad_usercomment,ad_userbot,ad_userlinks,s_blank
	ad_usertop=trim(request("ad_usertop"))
	ad_usercomment=trim(request("ad_usercomment"))
	ad_userbot=trim(request("ad_userbot"))
	ad_userlinks=trim(request("ad_userlinks"))
	if f_ext="shtml" or f_ext="asp" then
		s_blank=" "
	else
		s_blank=""
	end if
	set fso=server.createobject("scripting.filesystemobject")
	dirstr=server.MapPath("../ad/")
	if fso.FolderExists(dirstr)=false then fso.CreateFolder(dirstr)
	If ad_usertop<>"" and ad_usertop<>"　" Then	
		Call oblog.BuildFile(dirstr&"\ad_usertop.htm",ad_usertop)
		Call oblog.BuildFile(dirstr&"\ad_usertopjs.htm",oblog.htm2js(ad_usertop))
	Else
		Call oblog.BuildFile(dirstr&"\ad_usertop.htm",s_blank)
		Call oblog.BuildFile(dirstr&"\ad_usertopjs.htm",s_blank)
	End If
	If ad_usercomment<>"" and ad_usercomment<>"　" Then	
		Call oblog.BuildFile(dirstr&"\ad_usercomment.htm",ad_usercomment)
		Call oblog.BuildFile(dirstr&"\ad_usercommentjs.htm",oblog.htm2js(ad_usercomment))
	Else
		Call oblog.BuildFile(dirstr&"\ad_usercomment.htm",s_blank)
		Call oblog.BuildFile(dirstr&"\ad_usercommentjs.htm",s_blank)
	End If
	If ad_userbot<>"" and ad_userbot<>"　" Then	
		Call oblog.BuildFile(dirstr&"\ad_userbot.htm",ad_userbot)
		Call oblog.BuildFile(dirstr&"\ad_userbotjs.htm",oblog.htm2js(ad_userbot))
	Else
		Call oblog.BuildFile(dirstr&"\ad_userbot.htm",s_blank)
		Call oblog.BuildFile(dirstr&"\ad_userbotjs.htm",s_blank)
	End If
	If ad_userlinks<>"" and ad_userlinks<>"　" Then	
		Call oblog.BuildFile(dirstr&"\ad_userlinks.htm",ad_userlinks)
		Call oblog.BuildFile(dirstr&"\ad_userlinksjs.htm",oblog.htm2js(ad_userlinks))
	Else
		Call oblog.BuildFile(dirstr&"\ad_userlinks.htm",s_blank)
		Call oblog.BuildFile(dirstr&"\ad_userlinksjs.htm",s_blank)	
	End If	
	set fso=nothing
	oblog.showok "保存完成",""
else
	
%>
<html>
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css"> 
<body class="tdbg" style="background:#ffffff;">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr> 
    <td height="25" class="title"><div align="center"><strong><font color="#FFFFFF">修改用户blog页面广告</font></strong></div></td>
  </tr>
  <tr> 
    <td><form name="form1" method="post" action="admin_ad.asp">
	        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td>注意：3.1版本会自动将html代码转换为js代码，请直接在文本框内输入html即可。<br>
              <br>
              <br>
            用户页面顶部广告（此段代码显示在所有blog用户页面顶部，可以放置弹出广告等）
              <hr> 
                <textarea name="ad_usertop" cols="80" rows="8" ><%=oblog.readfile("../ad","ad_usertop.htm")%></textarea>
              <br>
              <br>
              用户页面底部广告（此段代码显示在所有blog用户页面最底部，可以书写版权信息等代码）
              <hr> 
                <textarea name="ad_userbot" cols="80" rows="8" ><%=oblog.readfile("../ad","ad_userbot.htm")%></textarea>
              <br>
              <br>
              用户回复上部广告（此段代码显示在所有blog用户评论窗口上部）
              <hr>
                <textarea name="ad_usercomment" cols="80" rows="8"><%=oblog.readfile("../ad","ad_usercomment.htm")%></textarea>
              <br>
              <br>
              用户友情连接部分广告（此段代码显示在所有blog用户友情连接标签部分）
              <hr> 
                <textarea name="ad_userlinks" cols="80" rows="8"><%=oblog.readfile("../ad","ad_userlinks.htm")%></textarea>
              <br>
              <br>
              <input name="Action" type="hidden" id="Action" value="saveconfig"> 
                <input type="submit" name="Submit" value="提交修改">
            </td>
          </tr>
        </table>
      </form></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
<%
end if
%>

