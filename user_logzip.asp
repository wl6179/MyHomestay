<!--#include file="conn.asp"-->
<!--#include file="inc/class_sys.asp"-->
<%
dim oblog
set oblog=new class_sys
oblog.start
if not oblog.checkuserlogined() then
	response.Redirect("login.asp")
end if
dim tablename,sql,filetype
dim rs,strLine
dim sdate,edate,nurl

filetype = lcase(trim(request("filetype")))
sdate=oblog.filt_badstr(request("selecty1")&"-"&request("selectm1")&"-"&request("selectd1"))
edate=oblog.filt_badstr(request("selecty2")&"-"&request("selectm2")&"-"&request("selectd2"))
if not isdate(sdate) then oblog.adderrstr("开始日期不正确")
if not isdate(edate) then oblog.adderrstr("结束日期不正确")
if oblog.errstr<>"" then
	oblog.showusererr
end if
nUrl=trim("http://" & Request.ServerVariables("SERVER_NAME"))
nUrl=lcase(nUrl & request.ServerVariables("SCRIPT_NAME"))
nurl=left(nUrl,instrrev(nUrl,"/"))
tablename = sdate&"到"&edate&"的文章"
if is_sqldata=1 then
	sql="select topic,addtime,logtext from oblog_log where userid="&oblog.logined_uid&" and addtime<='"&dateadd("d",1,edate)&"' and addtime>='"&sdate&"'"
else
	sql="select topic,addtime,logtext from oblog_log where userid="&oblog.logined_uid&" and addtime<=#"&dateadd("d",1,edate)&"# and addtime>=#"&sdate&"#"
end if
Set rs = oblog.Execute(sql)
if filetype="xml" then
	Response.contenttype="text/xml"
	Response.Charset = "gb2312"
	Response.AddHeader "Content-Disposition", "attachment;filename="&tablename&".xml"
	Response.write "<?xml version=""1.0"" encoding=""gb2312""?>"
	Response.write vbnewline&"<root>"
	strLine=""
	While not rs.EOF
		strLine= vbnewline&chr(9)&"<log>"
		strLine=  strLine &"<title>"&rs(0)&"</title>"
		strLine=  strLine & "<PubDate>"&rs(1)&"</PubDate>"
		strLine=  strLine &"<![CDATA["& newurl(rs(2),nurl) &"]]>"		
		strLine=  strLine &"</log>"	
		Response.write strLine 
		rs.MoveNext
	Wend
	Response.write vbnewline&"</root>"
elseif filetype="txt" then
	Response.contenttype="text"
	Response.AddHeader "Content-Disposition", "attachment;filename="&tablename&".txt"
	While not rs.EOF
		strLine=""
		strLine=strLine & "文章标题："&rs(0) & vbnewline
		strLine=strLine & "发布时间："&rs(1) & vbnewline
		strLine=strLine & "文章内容："&trim(stripHTML(rs(2)))
		Response.write strLine & vbnewline & vbnewline
		rs.MoveNext	
	Wend
else
if filetype="htm" then
		Response.contenttype="application/ms-download"
		Response.AddHeader "Content-Disposition", "attachment;filename="&tablename&".htm"
end if
	
	While not rs.EOF
	strLine= ""
	Response.write chr(9)&"<table style='font-size:10pt;TABLE-LAYOUT: fixed; WORD-BREAK: break-all' width='98%'align='center' bgColor=#ffffff border=1 >"& vbnewline
	Response.write chr(9)&"<tr>"& vbnewline
	strLine= strLine&chr(9)&chr(9)&"<td>"
	strLine= strLine&"文章标题："&rs(0)&"<br>"& vbnewline
	strLine= strLine&"发布时间："&rs(1)&"<br>"& vbnewline
	strLine= strLine&newurl(rs(2),nurl) &"</td>"& vbnewline
	Response.write strLine
	Response.write chr(9)&"</tr>"& vbnewline
	Response.write "</table><br>"& vbnewline
	rs.MoveNext
	Wend
end if
Set rs=nothing
Function stripHTML(strHTML)
  Dim objRegExp, strOutput
  Set objRegExp = New Regexp
  objRegExp.IgnoreCase = True
  objRegExp.Global = True
  objRegExp.Pattern = "<.+?>"
  strOutput = objRegExp.Replace(strHTML, "")
  strOutput = Replace(strOutput, "<", "<")
  strOutput = Replace(strOutput, ">", ">")
  stripHTML = replace(strOutput,"&nbsp;","")
  Set objRegExp = Nothing
End Function

Function newurl(strContent,byval url)           
    dim tempReg
    set tempReg=new RegExp
    tempReg.IgnoreCase=true
    tempReg.Global=true
    tempReg.Pattern="(^.*\/).*$"'含文件名的标准路径
    Url=tempReg.replace(url,"$1")
    tempReg.Pattern="((?:src|href).*?=[\'\u0022](?!ftp|http|https|mailto))"
    newurl=tempReg.replace(strContent,"$1"+Url)
    set tempReg=nothing
end Function
%>
