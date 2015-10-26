<%
from=lcase(Request.ServerVariables("HTTP_HOST")) 
if left(from,4)="www." or left(from,4)="myho" then
	Response.redirect("index.htm")'此处为网站首页地址
else
	'wangliang6179@163.com 王亮
	if left(from,8)="oldsite." then
		Response.redirect("http://www.myhomestay.com.cn/"& left(from,7) &"/index.htm")'特殊情况(针对oldsite文件夹)；
	elseif left(from,8)="newsite." then
		Response.redirect("http://www.myhomestay.com.cn/"& "oldsite/wangliang" &"/index.htm")'特殊情况(针对新版wl个人css网站的文件夹)；
	elseif left(from,4)="new." then
		Response.redirect("http://www.myhomestay.com.cn/"& "new_SiteModel" &"/index.html")	
		
	else
		Response.redirect("http://www.myhomestay.com.cn/"& left(from,3) &"/index.asp")'其他的，与‘www.’处理类似..WL
	end if
end if
%> 