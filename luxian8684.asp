<html>
<head>
<title>��������·��--385·--8684����������</title>
<meta name="keywords" content="��������·�� 385·">
<meta NAME="description" CONTENT="�ṩ�ͱ�������·��385·��ص���Ϣ��ѯ">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="/js/86841.css" type=text/css rel=stylesheet>
</head>
<Body>

<FORM action="http://beijing.8684.cn/so.php" method="get" name="form1" onSubmit="return LFormCheck()"; autocomplete="off">
<input type="hidden" name="cid" value="6" />
<p class="pv" align="center"><input id="pp" type=radio name=k value="pp" checked onClick='formonclick("1")';><label for="pp">����·�߲�ѯ</label><input id="ppp" type=radio name=k value=p  onClick='formonclick("2")';><label for="ppp">����վ���ѯ</label><input id="p2p" type=radio name=k value=p2p onClick='formonclick("3")';><label for="p2p">�������˲�ѯ</label></p>
<p align="center"><span class="pv" id="d"><input id=q maxlength=100 size=30 name=q value='385·'> <input type=submit value='��·��ѯ'></span></form></p><h1>��ҽ��Ժ·��ָ��</h1>

<%
Dim strLine, strLine2, codeLine, codeLine2

strLine = "323��"
strLine2 = "375·"

codeLine = Server.URLEncode(strLine)
codeLine2 = Server.URLEncode(strLine2)

Dim linkLine
linkLine = "http://beijing.8684.cn/pp_"
%>
ȥ��ҽ��Ժ��·�ߣ�
<a href="<% = linkLine & codeLine %>" target="_blank"><% = strLine %></a>
&nbsp;&nbsp;&nbsp;
<a href="<% = linkLine & codeLine2 %>" target="_blank"><% = strLine2 %></a>

<br><br>

<%
Dim strStation, strStation2, codeStation, codeStation2

strStation = "��ҽ��Ժ"
strStation2 = "�������պ����ѧ"

codeStation = Server.URLEncode(strStation)
codeStation2 = Server.URLEncode(strStation2)

Dim linkStation
linkStation = "http://beijing.8684.cn/so.php?cid=6&k=p&q="
%>
��ҽ��Ժ�ĳ�վ��
<a href="<% = linkStation & codeStation %>" target="_blank"><% = strStation %></a>
&nbsp;&nbsp;&nbsp;
<a href="<% = linkStation & codeStation2 %>" target="_blank"><% = strStation2 %></a>



</body>
</html>