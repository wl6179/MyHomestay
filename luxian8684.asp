<html>
<head>
<title>北京公交路线--385路--8684北京公交网</title>
<meta name="keywords" content="北京公交路线 385路">
<meta NAME="description" CONTENT="提供和北京公交路线385路相关的信息查询">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="/js/86841.css" type=text/css rel=stylesheet>
</head>
<Body>

<FORM action="http://beijing.8684.cn/so.php" method="get" name="form1" onSubmit="return LFormCheck()"; autocomplete="off">
<input type="hidden" name="cid" value="6" />
<p class="pv" align="center"><input id="pp" type=radio name=k value="pp" checked onClick='formonclick("1")';><label for="pp">公交路线查询</label><input id="ppp" type=radio name=k value=p  onClick='formonclick("2")';><label for="ppp">公交站点查询</label><input id="p2p" type=radio name=k value=p2p onClick='formonclick("3")';><label for="p2p">公交换乘查询</label></p>
<p align="center"><span class="pv" id="d"><input id=q maxlength=100 size=30 name=q value='385路'> <input type=submit value='线路查询'></span></form></p><h1>北医三院路线指南</h1>

<%
Dim strLine, strLine2, codeLine, codeLine2

strLine = "323快"
strLine2 = "375路"

codeLine = Server.URLEncode(strLine)
codeLine2 = Server.URLEncode(strLine2)

Dim linkLine
linkLine = "http://beijing.8684.cn/pp_"
%>
去北医三院的路线：
<a href="<% = linkLine & codeLine %>" target="_blank"><% = strLine %></a>
&nbsp;&nbsp;&nbsp;
<a href="<% = linkLine & codeLine2 %>" target="_blank"><% = strLine2 %></a>

<br><br>

<%
Dim strStation, strStation2, codeStation, codeStation2

strStation = "北医三院"
strStation2 = "北京航空航天大学"

codeStation = Server.URLEncode(strStation)
codeStation2 = Server.URLEncode(strStation2)

Dim linkStation
linkStation = "http://beijing.8684.cn/so.php?cid=6&k=p&q="
%>
北医三院的车站：
<a href="<% = linkStation & codeStation %>" target="_blank"><% = strStation %></a>
&nbsp;&nbsp;&nbsp;
<a href="<% = linkStation & codeStation2 %>" target="_blank"><% = strStation2 %></a>



</body>
</html>