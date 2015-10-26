<!--#include file="inc/inc_sys.asp"-->
<!--#include file="../inc/class_blog.asp"-->
<%

Dim Action,strYear,strMonth

Action=Lcase(trim(request("Action")))
If Action<>"hour" And Action<>"month" And Action<>"year" Then
	Response.Write "错误的操作参数"
	Response.End
End If
strYear=Request("year")
strMonth=Request("month")
strYear=CINT(strYear)
If strYear=0 Then strYear=Year(ServerDate(Date))
strMonth=CINT(strMonth)
If Not IsNumeric(strMonth) Then	strMonth=Month(ServerDate(Date))
%>
<html>
<head>
<title>系统分析工具</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
  <tr class="topbg"> 
    <td height="22" colspan=2 align=center><strong>系 统 分 析 工 具</strong></td>
  </tr>
  <form name="form1" action="admin_analyze.asp" method="get">
    <tr class="tdbg"> 
      <td width="100%" height="30">
      	<%=CreateSelectForm%>
       </td>
    </tr>
  </form>
   <tr class="tdbg"> 
      <td width="100%" height="30">
  <%
  	Dim rst,lngMax,intDays,MonthSql1,MonthSql2,strTable
  	strTable=" a"
	Select Case Action
		Case "year"			
			'注册增长
			Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(Userid) As iNum,Month(adddate) As iNode From oBlog_User Where Year(adddate)=" & strYear& " Group By Month(adddate))" & strTable)
			If rst.Eof Then
				lngMax=0
			Else
				If IsNull(rst(0)) Then
					lngMax=0
				Else	
					lngMax=rst(0)
				End If
			End If
			Set rst=oBlog.Execute("Select Count(Userid),Month(adddate) As iNode From oBlog_User Where Year(adddate)=" & strYear& " Group By Month(adddate)")
			If rst.Eof Then
				Set rst=Nothing
				Response.Write "没有相关数据"
				Response.End
			End If
			Call CreateChartTable(rst,12,JudgeRate(lngMax),strYear & "年月度注册增长趋势")
			'发帖增长
			Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(logid) As iNum,Month(addtime) As iNode From oBlog_log Where Year(addtime)=" & strYear& " Group By Month(addtime))" & strTable)
			If rst.Eof Then
				lngMax=0
			Else
				If IsNull(rst(0)) Then
					lngMax=0
				Else	
					lngMax=rst(0)
				End If
			End If
			Set rst=oBlog.Execute("Select Count(logid),Month(addtime) As iNode From oBlog_log Where Year(addtime)=" & strYear& " Group By Month(addtime)")
			If rst.Eof Then
				Set rst=Nothing
				Response.Write "没有相关数据"
				Response.End
			End If
			Call CreateChartTable(rst,12,JudgeRate(lngMax),strYear & "年月度发布增长趋势")
			'回复增长
			Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(commentid) As iNum,Month(addtime) As iNode From oBlog_comment Where Year(addtime)=" & strYear& " Group By Month(addtime))" & strTable)
			If rst.Eof Then
				lngMax=0
			Else
				If IsNull(rst(0)) Then
					lngMax=0
				Else	
					lngMax=rst(0)
				End If
			End If
			Set rst=oBlog.Execute("Select Count(commentid),Month(addtime) As iNode From oBlog_comment Where Year(addtime)=" & strYear& " Group By Month(addtime)")
			If rst.Eof Then
				Set rst=Nothing
				Response.Write "没有相关数据"
				Response.End
			End If
			Call CreateChartTable(rst,12,JudgeRate(lngMax),strYear & "年月度回复增长趋势")					
		Case "month"		
			If strMonth>0 Then
				intDays=MonthDays(strYear,strMonth)
				MonthSql1=" And Month(adddate)=" & strMonth
				MonthSql2=" And Month(addtime)=" & strMonth  
			Else
				intDays=31
			End If
			'注册增长
			Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(Userid) As iNum,Day(adddate) As iNode From oBlog_User Where Year(adddate)=" & strYear& MonthSql1 & " Group By Day(adddate))" & strTable)
			If rst.Eof Then
				lngMax=0
			Else
				If IsNull(rst(0)) Then
					lngMax=0
				Else	
					lngMax=rst(0)
				End If
			End If
			Set rst=oBlog.Execute("Select Count(Userid),Day(adddate) As iNode From oBlog_User Where Year(adddate)=" & strYear & MonthSql1 & " Group By Day(adddate)")
			If rst.Eof Then
				Set rst=Nothing
				Response.Write "没有相关数据"
				Response.End
			End If
			Call CreateChartTable(rst,intDays,JudgeRate(lngMax),strYear & "年" & strMonth & "月注册日增长趋势")
			'发帖增长
			Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(logid) As iNum,Day(addtime) As iNode From oBlog_log Where Year(addtime)=" & strYear & MonthSql2 & " Group By Day(addtime))" & strTable)
			If rst.Eof Then
				lngMax=0
			Else
				If IsNull(rst(0)) Then
					lngMax=0
				Else	
					lngMax=rst(0)
				End If
			End If
			Set rst=oBlog.Execute("Select Count(logid),Day(addtime) As iNode From oBlog_log Where Year(addtime)=" & strYear& MonthSql2 & " Group By Day(addtime)")
			If rst.Eof Then
				Set rst=Nothing
				Response.Write "没有相关数据"
				Response.End
			End If
			Call CreateChartTable(rst,intDays,JudgeRate(lngMax),strYear & "年" & strMonth & "月发布日增长趋势")
			'回复增长
			Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(commentid) As iNum,Day(addtime) As iNode From oBlog_comment Where Year(addtime)=" & strYear & MonthSql2 & " Group By Day(addtime))" & strTable)
			If rst.Eof Then
				lngMax=0
			Else
				If IsNull(rst(0)) Then
					lngMax=0
				Else	
					lngMax=rst(0)
				End If
			End If
			Set rst=oBlog.Execute("Select Count(commentid),day(addtime) As iNode From oBlog_comment Where Year(addtime)=" & strYear& MonthSql2 & " Group By day(addtime)")
			If rst.Eof Then
				Set rst=Nothing
				Response.Write "没有相关数据"
				Response.End
			End If
			Call CreateChartTable(rst,intDays,JudgeRate(lngMax),strYear & "年" & strMonth & "月回复日增长趋势")		
		Case "hour"
			If strMonth>0 Then
				MonthSql1=" And Month(adddate)=" & strMonth
				MonthSql2=" And Month(addtime)=" & strMonth   
			End If
			'注册增长
			If Is_Sqldata=1 Then
				Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(Userid) As iNum,{ fn HOUR(adddate) } As iNode From oBlog_User Where Year(adddate)=" & strYear& MonthSql1 & " Group By { fn HOUR(adddate) })" & strTable)
			Else
				Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(Userid) As iNum,Hour(adddate) As iNode From oBlog_User Where Year(adddate)=" & strYear& MonthSql1 & " Group By Hour(adddate))" & strTable)
			End If				
			
			If rst.Eof Then
				lngMax=0
			Else
				If IsNull(rst(0)) Then
					lngMax=0
				Else	
					lngMax=rst(0)
				End If
			End If
			If Is_SqlData=1 Then
				Set rst=oBlog.Execute("Select Count(Userid),{ fn HOUR(adddate) } As iNode From oBlog_User Where Year(adddate)=" & strYear& MonthSql1 &  " Group By { fn HOUR(adddate) }")
			Else
				Set rst=oBlog.Execute("Select Count(Userid),Hour(adddate) As iNode From oBlog_User Where Year(adddate)=" & strYear& MonthSql1 &  " Group By Hour(adddate)")
			End If
			If rst.Eof Then
				Set rst=Nothing
				Response.Write "没有相关数据"
				Response.End
			End If
			Call CreateChartTable(rst,24,JudgeRate(lngMax),strYear & "年" & strMonth & "月时段注册增长趋势")
			'发帖增长			
			If Is_SqlData=1 Then
				Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(logid) As iNum,{ fn HOUR(addtime) } As iNode From oBlog_log Where Year(addtime)=" & strYear&  MonthSql2 & " Group By { fn HOUR(addtime) })" & strTable)
			Else
				Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(logid) As iNum,Hour(addtime) As iNode From oBlog_log Where Year(addtime)=" & strYear&  MonthSql2 & " Group By Hour(addtime))" & strTable)
			End If
			If rst.Eof Then
				lngMax=0
			Else
				If IsNull(rst(0)) Then
					lngMax=0
				Else	
					lngMax=rst(0)
				End If
			End If			
			If Is_SqlData=1 Then
				Set rst=oBlog.Execute("Select Count(logid),{ fn HOUR(addtime) } As iNode From oBlog_log Where Year(addtime)=" & strYear& MonthSql2 & " Group By { fn HOUR(addtime) }")
			Else
				Set rst=oBlog.Execute("Select Count(logid),Hour(addtime) As iNode From oBlog_log Where Year(addtime)=" & strYear& MonthSql2 & " Group By Hour(addtime)")
			End If
			If rst.Eof Then
				Set rst=Nothing
				Response.Write "没有相关数据"
				Response.End
			End If
			Call CreateChartTable(rst,24,JudgeRate(lngMax),strYear & "年" & strMonth & "月时段发布增长趋势")
			'回复增长
			If Is_SqlData=1 Then
				Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(commentid) As iNum,{ fn HOUR(addtime) } As iNode From oBlog_comment Where Year(addtime)=" & strYear& MonthSql2 & " Group By { fn HOUR(addtime) })" & strTable)
			Else
				Set rst=oBlog.Execute("Select Max(iNum) From (Select Count(commentid) As iNum,Hour(addtime) As iNode From oBlog_comment Where Year(addtime)=" & strYear& MonthSql2 & " Group By Hour(addtime))" & strTable)
			End If
			If rst.Eof Then
				lngMax=0
			Else
				If IsNull(rst(0)) Then
					lngMax=0
				Else	
					lngMax=rst(0)
				End If
			End If
			If Is_SqlData=1 Then
				Set rst=oBlog.Execute("Select Count(commentid),{ fn HOUR(addtime) } As iNode From oBlog_comment Where Year(addtime)=" & strYear& MonthSql2 & " Group By { fn HOUR(addtime) }")
			Else
				Set rst=oBlog.Execute("Select Count(commentid),Hour(addtime) As iNode From oBlog_comment Where Year(addtime)=" & strYear& MonthSql2 & " Group By Hour(addtime)")
			End If
			If rst.Eof Then
				Set rst=Nothing
				Response.Write "没有相关数据"
				Response.End
			End If
			Call CreateChartTable(rst,24,JudgeRate(lngMax),strYear & "年" & strMonth & "月时段回复增长趋势")	
 	End Select
 	Set rst=Nothing
 	Set oBlog =Nothing
  %>
  </td>
    </tr>
</table>

<%


'创建日期下拉条
'如果是年度分析,只取年度
'如果是月度分析,则只取年月
'如果是时段分析,则不必显示该Form
Function CreateSelectForm()	
	Dim SubmitStr,YearStr,MonthStr,intYear,i
	SubmitStr="<input type=hidden name=action value="&Action&"><input type=submit value=分析>"
	intYear = Year(ServerDate(Date))
	YearStr="<Select name=year>" & Vbcrlf
	For i=intYear To intYear-5 Step -1
		YearStr= YearStr & "<option Value=" & i &">" & i & "</option>" & Vbcrlf
	Next
	YearStr= YearStr & "</Select>年" & Vbcrlf
	If Action="year" Then 
		CreateSelectForm= "&nbsp;&nbsp;分析区段：" & YearStr & SubmitStr
		Exit Function
	End If
	MonthStr="<Select name=month>" & Vbcrlf
	For i=1 To 12
		MonthStr= MonthStr & "<option Value=" & i &">" & i & "</option>" & Vbcrlf
	Next
	MonthStr= MonthStr & "<option Value=0>全部</option>" & Vbcrlf
	MonthStr= MonthStr & "</Select>月"
	CreateSelectForm=  "&nbsp;&nbsp;分析区段：" & YearStr & MonthStr& SubmitStr
End Function

'通用图表创建程序
Function CreateChartTable(rst,intNodes,intRate,strTitle)
	Dim i,lngNode,BarString,NumString
	'Dim intRate '比率换算，取最大值，如果最大值小于100，则比率为1；1000则为0.1;依次压缩
	'特殊处理0月
	strTitle=Replace(strTitle,"年0月","年度")
	Response.Write "<p>&nbsp;</p>" & Vbcrlf
	Response.Write "<table style=""border:1px #EEEEEE dotted;font-size:12px"" align=center>"
	Response.Write "<tr align=center colspan=31>"& strTitle &"</tr>" & VbCrlf	
	For i=1 to intNodes
		rst.Filter="inode=" & i
		If rst.Eof Then
			lngNode=0
		Else
			If IsNull(rst(0)) Then
				lngNode=0
			Else
				lngNode=rst(0)
			End If
		End If
		BarString = BarString & "<td width=15 valign=baseline><img src=images/bar2.gif width=12 height=" & lngNode*intRate & " title=" & lngNode & "></td>" & VbCrlf
		NumString = NumString & "<td>" & i & "</td>" & VbCrlf
	Next
	Response.Write "<tr>" & BarString & "</tr>" & VbCrlf
	Response.Write "<tr>" & NumString & "</tr>" & VbCrlf
	Response.Write "</table>"& VbCrlf
	Response.Write "<p>&nbsp;</p>" & VbCrlf
End Function

'判断收缩比率
Function JudgeRate(lngNum)
	If lngNum<=100 Then
		JudgeRate=1
	ElseIf lngNum>100 And lngNum<=1000 Then
		JudgeRate=0.1
	ElseIf lngNum>1000 And lngNum<=10000 Then
		JudgeRate=0.01
	ElseIf lngNum>10000 And lngNum<=100000 Then
		JudgeRate=0.001
	ElseIf lngNum>100000 And lngNum<=1000000 Then
		JudgeRate=0.0001
	End If
End Function

Function MonthDays(intYear,intMonth)
	Dim strDate1,strDate2
	strDate1= CDate(intYear & "-" & intMonth & "-1")
	strDate2=DateAdd("m",1,CDate(strDate1))
	MonthDays=Datediff("d",strDate1,strDate2)
End Function
%>