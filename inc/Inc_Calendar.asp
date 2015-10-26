<%
'*********************************************************
'File: 			Inc_Calendar.asp
'Description:	日历函数模块 For oBlog3.1
'Author: 		王旭 
'Copyright:		http://www.myhomestay.com.cn
'LastUpdate:	20050926
'*********************************************************

'------------------------------------------------
'Calendar(intYear,intMonth,intUserId) 
'intYear:年
'intMonth:月
'intUserId:  用户编号，如果为空则为系统操作
'------------------------------------------------
Function Calendar(intYear,intMonth,intUserId) 
	Dim strContent 
	Dim MonthName(12), MonthDays(12)
	ReDim aDays(2,0)
	Dim intCount,strSql
	Dim thisYear,thisMonth,thisDay, nextYear,nextMonth,preYear,preMonth,theDay,startWeek,toDay,toMonth,toYear
	Dim rst,Link_TF	 
	Dim i,j,k,l,m
	Dim CommondFile,strUserSql
	CommondFile= "MyHomestay."&f_ext&"?uid="&intUserId&"&do=month&month="
	intUserId=CheckInt(intUserId)
	If intUserId="" OR  intUserId=0 Then
		strUserSql=" "
		CommondFile=blogdir & CommondFile
	Else
		strUserSql=" And userid=" & intUserId & " "
	End If
	intCount=0
	
	IF intYear=Empty Then intYear=Year(Now())
	IF intMonth=Empty Then intMonth=Month(Now())
	
	intYear=CInt(intYear)
	intMonth=CInt(intMonth)

	
	thisYear=intYear
	thisMonth=intMonth
	thisDay=0
	
	toDay=CInt(Day(Now()))
	toMonth=CInt(Month(Now()))
	toYear=CInt(Year(Now()))
	
	MonthName(0)=""
	MonthName(1)="1"
	MonthName(2)="2"
	MonthName(3)="3"
	MonthName(4)="4"
	MonthName(5)="5"
	MonthName(6)="6"
	MonthName(7)="7"
	MonthName(8)="8"
	MonthName(9)="9"
	MonthName(10)="10"
	MonthName(11)="11"
	MonthName(12)="12"
	
	
	MonthDays(0)=""
	MonthDays(1)=31
	MonthDays(2)=28
	MonthDays(3)=31
	MonthDays(4)=30
	MonthDays(5)=31
	MonthDays(6)=30
	MonthDays(7)=31
	MonthDays(8)=31
	MonthDays(9)=30
	MonthDays(10)=31
	MonthDays(11)=30
	MonthDays(12)=31
	
	'读取一次数据库,只处理掉草稿即可
	strSql="SELECT Count(logid) as lognum, Day(Addtime) as logday FROM oblog_log WHERE Year(addtime)=" & intYear & " AND Month(addtime)=" & intMonth & " And isdraft=0 " & strUserSql & " Group By Day(Addtime)"
	Set rst=Server.CreateObject("ADODB.RecordSet")
	'Response.Write strSql
	'If Not IsObject(conn) Then Link_DataBase
	rst.Open strSql,Conn,1,3
	'Set rst=oBlog.Execute(strSql)	
	'设置有文章的数组
	theDay=0
	Do While Not rst.EOF
		IF rst("logday")<>theDay Then
			theDay=rst("logday")
			ReDim PreServe aDays(2,intCount)
			aDays(0,intCount)=intMonth
			aDays(1,intCount)=theDay
			aDays(2,intCount)="MyHomestay."&f_ext&"?uid="&intUserId&"&do=day&day=" & CStr(CDate(intYear & "-" & intMonth & "-" & theDay))
			intCount=intCount+1
		End IF
		rst.MoveNext
	Loop
	rst.Close
	Set rst=Nothing			
	
	'处理闰年
	If IsDate("February 29, " & thisYear) Then MonthDays(2)=29	
	
	startWeek=WeekDay(intMonth&"-1-"&intYear)-1
	
	'生成前一月后一月标志
	nextMonth=intMonth+1
	nextYear=intYear
	IF nextMonth>12 then 
		nextMonth=1
		nextYear=nextYear+1
	End IF
	preMonth=intMonth-1
	preYear=intYear
	IF preMonth<1 then 
		preMonth=12
		preYear=preYear-1
	End IF
	strContent=""
	strContent=strContent & ("<table width=""100%"" >")
	strContent=strContent & ("<caption><a href="""& CommondFile & (intYear-1) & Right("0" & intMonth,2) &""" title=""上一年""><span class=""arrow""><<</span></a>&nbsp;&nbsp;<a href=""" & CommondFile & preYear& Right("0" & preMonth,2)&""" title=""上一月""><span class=""arrow""><</span></a>&nbsp;"&intYear&"<a href=""" & CommondFile & Year(Date) & Right("0" & Month(Date),2) & """ title=""返回当月""> - </a>"& MonthName(intMonth)&"&nbsp;<a href="""& CommondFile & nextYear& Right("0" & nextMonth,2) &""" title=""下一月""><span class=""arrow"">></span></a>&nbsp;&nbsp;<a href=""" & CommondFile & intYear+1 & Right("0" & intMonth,2) &""" title=""下一年""><span class=""arrow"">>></span></a></caption><tr>")
	'strContent=strContent & ("<caption><a href=""" & CommondFile & preYear& Right("0" & preMonth,2)&""" title=""上一月""><span class=""arrow""><</span></a> "&intYear&"  <a href=""" & CommondFile & Year(Date) & Right("0" & Month(Date),2) & """ title=""返回当月"">-</a> "& MonthName(intMonth)&" <a href="""& CommondFile & nextYear& Right("0" & nextMonth,2) &""" title=""下一月""><span class=""arrow"">></span></a></caption><tr>")
	strContent=strContent & ("<th>日</th><th>一</th><th>二</th><th>三</th><th>四</th><th>五</th><th>六</th></tr><tr>")
	
	For  i=0 TO startWeek-1
		strContent=strContent & ("<td align=""center"">&nbsp;</td>")
	Next
	
	j=1
	While j<=MonthDays(thisMonth)
	 	For k=startWeek To 6
			strContent=strContent & ("<td align=""center"">")
			Link_TF="Flase"
			For l=0 TO Ubound(aDays,2)
				IF aDays(0,l)<>"" Then
					IF aDays(0,l)=thisMonth AND aDays(1,l)=j Then
						strContent=strContent & ("<a href="""&aDays(2,l)&""">")
						Link_TF="True"
					End IF
				End IF
			Next
		IF j<=MonthDays(thisMonth) Then strContent=strContent & (j)
		IF Link_TF="True" Then strContent=strContent & ("</a>")
        strContent=strContent & ("</td>")
		j=j+1
	Next
	startWeek=0
	strContent=strContent & ("</tr>")
	Wend
	strContent=strContent & ("</table>")
	Calendar=strContent 
	strContent=""
End Function
%>