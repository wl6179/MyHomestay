<%
from=lcase(Request.ServerVariables("HTTP_HOST")) 
if left(from,4)="www." or left(from,4)="myho" then
	Response.redirect("index.htm")'�˴�Ϊ��վ��ҳ��ַ
else
	'wangliang6179@163.com ����
	if left(from,8)="oldsite." then
		Response.redirect("http://www.myhomestay.com.cn/"& left(from,7) &"/index.htm")'�������(���oldsite�ļ���)��
	elseif left(from,8)="newsite." then
		Response.redirect("http://www.myhomestay.com.cn/"& "oldsite/wangliang" &"/index.htm")'�������(����°�wl����css��վ���ļ���)��
	elseif left(from,4)="new." then
		Response.redirect("http://www.myhomestay.com.cn/"& "new_SiteModel" &"/index.html")	
		
	else
		Response.redirect("http://www.myhomestay.com.cn/"& left(from,3) &"/index.asp")'�����ģ��롮www.����������..WL
	end if
end if
%> 