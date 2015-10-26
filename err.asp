<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
dim show_err,errmsg,strerr,errmsg1,i
errmsg=trim(request("message"))
call sysshow()
If errmsg<>"" Then
	errmsg=Split(oblog.filt_html(errmsg),"_")
	For i=0 to UBound(errmsg)
		If i=0 Then
			errmsg1=errmsg1&"<li>"&errmsg(i)
		Else
			errmsg1=errmsg1&"<br><li>"&errmsg(i)
		End If
	Next 
End If 
show_err= "<table cellpadding=2 cellspacing=1 border=0 width=400 align=center>" & vbcrlf
show_err=show_err & "  <br><tr align='center'><td height='22' ><strong>错误信息</strong><hr noshade></td></tr>" & vbcrlf
show_err=show_err & "  <tr><td height='100'  valign='top'><b>产生错误的可能原因：</b><br><font color=red>" & errmsg1 &"</font></td></tr>" & vbcrlf
show_err=show_err & "  <tr align='center'><td ><input type='button' name='historybackwl' value='返回上一页' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></td></tr>" & vbcrlf
show_err=show_err & "</table>" & vbcrlf
show=replace(show,"$show_list$",show_err)&oblog.setup(7,0)&"homestay"
response.Write show
%>