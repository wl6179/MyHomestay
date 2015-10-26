<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<%dim oblog
set oblog=new class_sys
oblog.autoupdate=false
oblog.start
Response.contentType="application/xml"
Response.Expires=0
Response.Write("<?xml version=""1.0"" encoding=""gb2312""?>")%>
<rss version="2.0">
<channel>
<title><%=FmtStringForXML(oblog.setup(1,0))%></title>
<link><%=oblog.setup(3,0)%></link>
<description><%=oblog.setup(5,0)%></description>
<generator>Oblog 3.0</generator>
<webMaster><%=oblog.setup(6,0)%></webMaster>
<%
dim rs,sql,classid,logtext
classid=clng(request("classid"))
if classid>0 then
	sql=" and classid="&classid
else
	sql=""
end if
set rs=oblog.execute("select top 12 * from oblog_log where ishide=0 and passcheck=1 and isdraft=0 and blog_password=0"&sql&" order by logid desc")
if rs.Eof or rs.Bof then
  response.write "<item></item>"
end if
while not rs.Eof 
	if rs("ispassword")="" or isnull(rs("ispassword")) then
		logtext=oblog.trueurl(rs("logtext"))
	else
		logtext="此文章内容已加密"
	end if
	if rs("ishide")=1 then logtext="此文章内容已隐藏"
    response.Write "<item>" & vbcrlf
	Response.write "<title><![CDATA["&rs("topic")&"]]></title>" & vbcrlf
	if true_domain=0 then
		response.write "<link>"&oblog.setup(3,0)&rs("logfile")&"</link>" & vbcrlf
	else
		response.write "<link>"&rs("logfile")&"</link>" & vbcrlf
	end if
	response.write "<author>"&rs("author")&"</author>" & vbcrlf
	response.write "<pubDate>"&rs("addtime")&"</pubDate>" & vbcrlf
 	response.write "<description><![CDATA["&logtext&"]]></description>" & vbcrlf
	response.write "</item>"
	rs.MoveNext          
wend                  
set rs=nothing
%>
</channel>
</rss>
<%
Function FmtStringForXML(byval sContent)
	Dim objRegExp,strOutput
	If IsNull(sContent) Then
		FmtStringForXML=""
		Exit Function
	End If
	strOutput=Trim(sContent)
	If Instr(strOutput,"<") And Instr(strOutput,">")  Then
		'剔除<>标记
		Set objRegExp = New Regexp
		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		objRegExp.Pattern = "<.+?>"
		strOutput = objRegExp.Replace(strOutput, "")
		Set objRegExp = Nothing	
	End If
	strOutput=Replace(strOutput," ","")
	strOutput=Replace(strOutput,"<>","")
	strOutput=Replace(strOutput,"&nbsp;"," ")
	FmtStringForXML = strOutput	
End Function
%>