<%
'*********************************************************
'File: 			Inc_Functions.asp
'Description:	公用函数模块 For 3.1
'Author: 		王旭 
'Copyright:		http://www.myhomestay.com.cn
'LastUpdate:	20050921
'*********************************************************

'	ProIndexLog
'输入参数：
'	strHTML Logtext内容
'	iLen  截取的长度，以HTML代码为主，主要是为了后面f的分行显示
'	iLine 默认显示行数
Public Function ProIndexLog(strHTML,iLen,iLine) 
	ON ERROR RESUME NEXT
	Dim aHTML,i,j,lP1,lP2	
	Dim objRegExp, strOutput,sOut
	
	If Instr(strHTML,"#此前在首页部分显示#") Then
		strOutput=Split(strHTML,"#此前在首页部分显示#")(0) 
	Else
		strHTML=Left(strHTML,iLen)
		lP1=InStrRev(strHTML,"<")
		lP2=InStrRev(strHTML,">")
		'取500文本，并且判断最后一个<出现的位置，如果该<之后没有出现>
		'则补充一个进行闭合
		If lP1>0 Then
			If lP2<lP1 Then
				strHTML=strHTML & ">"
			End If
		End If
		'确认断行位置
		strHTML=Replace(strHTML,"<BR>","$breakline$")
		strHTML=Replace(strHTML,"<BR/>","$breakline$")
		strHTML=Replace(strHTML,"<br/>","$breakline$")
		strHTML=Replace(strHTML,"<br />","$breakline$")
		strHTML=Replace(strHTML,"<br>","$breakline$")
		strHTML=Replace(strHTML,"</p>","$breakline$")
		strHTML=Replace(strHTML,"</P>","$breakline$")
		'剔除<>标记
		Set objRegExp = New Regexp'建立正则表达式；
		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		objRegExp.Pattern = "<.+?>"
		strOutput = objRegExp.Replace(strHTML, "")
		Set objRegExp = Nothing	
		'以$breakline$分割，取iLine行
		aHTML=Split(strOutput,"$breakline$")
		j=0
		For i=0 to Ubound(aHTML)
			j=j+1
			strOutput=	Trim(aHTML(i))
			If strOutput<>"" Then
				If j<iLine then
					sOut=sOut & strOutput & "<BR/>" & Vbcrlf
				else			
					'在第iLine行处进行判断,如果该行大于10个字则显示，否则不显示该行
					'防止该行破坏显示结果
					If j=iLine Then 
						If Len(strOutput)>10 Then 
							sOut=sOut & strOutput & "<BR/>" & Vbcrlf
						Else
							'sOut=sOut & ""
						End If					
					else
						Exit For
					end if
				end if
			Else
				sOut=sOut & strOutput & "<BR/>" & Vbcrlf
			End If 			
		Next
		strOutput=Replace(sOut,"<BR/><BR/>","<BR/>")		
		strOutput=Replace(sOut,"。","。<BR/>")		
		If UCase(Right(strOutput,4))="<BR>" Then
			strOutput = Left(strOutput, Len(strOutput)-4)
		ElseIf UCase(Right(strOutput,5))="<BR/>" Then
			strOutput = Left(strOutput, Len(strOutput)-5)
		End If
	End If
	'文字之间进行换行标记此处写在函数外部			
	'ProIndexLog=  strOutput & " ..." & "<HR SIZE=1 color=gray />"
	ProIndexLog=  strOutput & " ..."
End Function 


function unHtml(content)
	ON ERROR RESUME NEXT
	unHtml=content
	if content <> "" then
		unHtml=Server.HTMLEncode(unHtml)
		unHtml=replace(unHtml,vbcrlf,"<br>")
		unHtml=replace(unHtml,chr(9),"&nbsp;&nbsp;&nbsp;&nbsp;")
		unHtml=replace(unHtml," ","&nbsp;")
		unHtml=replace(unHtml,"&","")
		unHtml=replace(unHtml,"?","")
	end if
end function

'------------------------------------------------
'EncodeJP(byval strContent)
'日文编码
'10k文章编码过程小于0.01秒，不会影响到执行效率
'目前需要更新的位置为：
'站点配置里的各个项目：名称、描述
'发布文章时的标题、内容、关键字
'发布留言/评论时的内容
'搜索时对关键字进行编码
'暂时不考虑注册名问题
'可与其他函数配合使用
'------------------------------------------------
Function EncodeJP(byval strContent)

	If strContent="" Then Exit Function
	
	'SQL版本不进行编码
	If IS_SQLDATA=1 Then 
		EncodeJP=strContent
		Exit Function
	End If
	
	strContent=Replace(strContent,"ガ","&#12460;")
    strContent=Replace(strContent,"ギ","&#12462;")
    strContent=Replace(strContent,"グ","&#12464;")
    strContent=Replace(strContent,"ア","&#12450;")
    strContent=Replace(strContent,"ゲ","&#12466;")
    strContent=Replace(strContent,"ゴ","&#12468;")
    strContent=Replace(strContent,"ザ","&#12470;")
    strContent=Replace(strContent,"ジ","&#12472;")
    strContent=Replace(strContent,"ズ","&#12474;")
    strContent=Replace(strContent,"ゼ","&#12476;")
    strContent=Replace(strContent,"ゾ","&#12478;")
    strContent=Replace(strContent,"ダ","&#12480;")
    strContent=Replace(strContent,"ヂ","&#12482;")
    strContent=Replace(strContent,"ヅ","&#12485;")
    strContent=Replace(strContent,"デ","&#12487;")
    strContent=Replace(strContent,"ド","&#12489;")
    strContent=Replace(strContent,"バ","&#12496;")
    strContent=Replace(strContent,"パ","&#12497;")
    strContent=Replace(strContent,"ビ","&#12499;")
    strContent=Replace(strContent,"ピ","&#12500;")
    strContent=Replace(strContent,"ブ","&#12502;")
    strContent=Replace(strContent,"ブ","&#12502;")
    strContent=Replace(strContent,"プ","&#12503;")
    strContent=Replace(strContent,"ベ","&#12505;")
    strContent=Replace(strContent,"ペ","&#12506;")
    strContent=Replace(strContent,"ボ","&#12508;")
    strContent=Replace(strContent,"ポ","&#12509;")
    strContent=Replace(strContent,"ヴ","&#12532;")

    EncodeJP=strContent
End Function

'------------------------------------------------
'Pause(byval iCount)
'暂停功能
'用于批量转移，转换，生成过程中，防止持续耗费系统资源
'------------------------------------------------
Sub Pause()
	Dim i,lStep,iCount
	iCount=P_BLOG_UPDATEPAUSE
	'本机测试执行时间为0.03~0.05
	lStep=200000	
	'如果为0或者非数值则不限制
	If  Not IsNumeric(iCount) OR iCount=0 Then Exit Sub
	iCount=CLNG(iCount)
	'Response.Write timer & "<br>"
	'本机测试3~5秒
	If iCount>100 Then iCount=100
	For i=1 To iCount * lStep
	Next
	'Response.Write timer 
End Sub

'------------------------------------------------
'CheckValidEnName(byval strName)检测有效的英文ID注册帐号，为英文字母和数字和下划线！
'只允许数字(48~57)+大(65~90)小(97~122)写字母和下划线
'------------------------------------------------
Function CheckValidEnName(byval strName)
	Dim objReg
	CheckValidEnName = True
	
	If IsNull(strName) OR strName="" Then Exit Function
		
	Set objReg=New RegExp
	objReg.IgnoreCase =True
	objReg.Global=True	
	objReg.Pattern="^\w+$"
	CheckValidEnName = objReg.Test(strName)
	'如果不允许全部为数字的ID，wl
	If en_nameisnum=0 Then
		If IsNumeric(strName) Then 	CheckValidEnName=False		
	End If
	Set objReg=Nothing
End Function

'------------------------------------------------
'FilterJS(strHTML)
'过滤脚本
'------------------------------------------------
Function FilterJS(byval strHTML)
	Dim objReg,strContent		
	If IsNull(strHTML) OR strHTML="" Then Exit Function		
			
	Set objReg=New RegExp
	objReg.IgnoreCase =True
	objReg.Global=True
	objReg.Pattern="(&#)"
	strContent=objReg.Replace(strHTML,"")
	objReg.Pattern="(function|meta|value|window\.|script|js:|about:|file:|Document\.|vbs:|frame|cookie)"
	strContent=objReg.Replace(strContent,"")
	objReg.Pattern="(on(finish|mouse|Exit=|error|click|key|load|focus|Blur))"
	strContent=objReg.Replace(strContent,"")
	FilterJS=strContent
	strContent=""
	Set objReg=Nothing		
End Function

'------------------------------------------------
'CheckInt(byval strNumber)
'检查并转换整形值
'------------------------------------------------
Function CheckInt(byval strNumber)
	If isNull(strNumber) OR Not IsNumeric(strNumber) Then
		CheckInt=""		
	Else
		CheckInt=CLNG(strNumber)
	End If
End Function

'------------------------------------------------
'ProtectSql(sSql)
'用于接收地址栏参数传递时SQL组合保护
'------------------------------------------------
'防止SQL注入
Function ProtectSQL(sSql)
	If ISNull(sSql) Then Exit Function
	sSql=Trim(sSql)
	If sSql="" Then Exit Function 
	sSql=Replace(sSql,Chr(0),"")	
	sSql=Replace(sSql,"'","‘")
	sSql=Replace(sSql," ","")
	sSql=Replace(sSql,"%","％")
	sSql=Replace(sSql,"-","－")		
	ProtectSQL=sSql
End Function

'用于用户发布的各种信息过滤，带脏话过滤
Function HTMLEncode(fString)
	If Not IsNull(fString) Then
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
		fString = Replace(fString, CHR(32), " ")		'&nbsp;
		fString = Replace(fString, CHR(9), " ")			'&nbsp;
		fString = Replace(fString, CHR(34), "&quot;")
		'fString = Replace(fString, CHR(39), "&#39;")	'单引号过滤
		fString = Replace(fString, CHR(13), "")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<BR> ")
		'fString=ChkBadWords(fString)
		HTMLEncode = fString
	End If
End Function

'------------------------------------------------	
'RemoveHtml(byval strConten
'移除HTML标记
'主要用户保存到数据库前的过滤
'------------------------------------------------	
Function RemoveHtml(byval strContent)
	Dim objReg ,strTmp
	If strContent="" OR ISNull(strContent) Then Exit Function
		
	Set objReg=new RegExp
	objReg.IgnoreCase =True
	objReg.Global=True
	objReg.Pattern="<(.[^>]*)>"
	strTmp=objReg.Replace(strContent, "")
	Set objReg=Nothing
	RemoveHtml=strTmp
	strTmp=""
End Function
	
'------------------------------------------------
'RedirectBy301(strURL)
'针对搜索引擎进行301重定向，立即更新目标地址
'------------------------------------------------
Sub RedirectBy301(byval strURL)
	Response.Clear
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location",strURL
	Response.End
End Sub

'------------------------------------------------
'ServerDate(byval strDate)
'服务器时差设置
'回复/留言及发布文章
'接收Trackback
'------------------------------------------------
Function ServerDate(byval strDate)
	Dim intHours
	If Not isDate(strDate) Then Exit Function
	intHours=P_Site_Hours
	If Not isNumeric(intHours) Then 
		intHours=0
		ServerDate=strDate
		Exit Function
	End If
	intHours =CLNG(intHours)
	If intHours>24 OR intHours<-24 Then 
		intHours=0
		ServerDate=strDate
		Exit Function
	End If
	ServerDate=Dateadd("h",intHours,strDate)	
End Function

'经测试使用此方法比include方法还要慢
Function ReadFileToString(byval oFSO,byval userpath,byval sFile)
'对目录进行处理
'该文件是从最底部开始的
'On Error Resume Next
	Dim oStream
	'处理最顶层的inc
	sFile=Replace(sFile,"..\..\..\..\","")	
	sFile=Replace(sFile,"..\inc\",userpath & "\inc\")
	sFile=Replace(sFile,"calendar\",userpath & "\calendar\")
	sFile=Replace(sFile,"subject\",userpath & "\subject\")
	sFile=Replace(sFile,"archives\",userpath & "\archives\")	
	sFile=Replace(sFile,"\\","\")
	sFile=Replace(sFile,"..\","")
	sFile=Replace(sFile,"\","/")
	sFile=Replace(sFile,"..","")
	'Response.Write "sFile:" & sFile
	'此处暂时不必判断文件是否真实存在
	Set oStream=oFSO.OpenTextFile(Server.Mappath(sFile),1,False)
	ReadFileToString = oStream.ReadAll
	Set oStream=Nothing
	'If Err.Number>0 Then ReadFileToString=""
End Function

'获取访问者IP
Function GetIP()

    Dim strIPAddr
    If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then
        strIPAddr = Request.ServerVariables("REMOTE_ADDR")
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then
        strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)
    ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then
        strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
    Else
        strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    End If
    GetIP = ProtectSQL(Trim(Mid(strIPAddr, 1, 30)))
End Function

Function FilterIDs(byval strIDs)
	Dim arrIDs,i,strReturn
	strIDs=Trim(strIDs)
	If Len(strIDs)=0  Then Exit Function
	arrIDs=Split(strIDs,",")
	For i=0 To Ubound(arrIds)
		If IsNumeric(arrIDs(i)) Then
			strReturn=strReturn & "," & CLNG(arrIDs(i))
		End If	
	Next
	If Left(strReturn,1)="," Then strReturn=Right(strReturn,Len(strReturn)-1)
	FilterIDs=strReturn	
End Function
	
Sub SystemState()
	If Application(cache_name_user&"_systemstate")="stop"  Then
		If Session("adminname")="" Then
			If Right(LCase(Request.ServerVariables("SCRIPT_NAME")),16)<>"/administratorLogin.asp" Then
	%>
	<style type="text/css">
		.border
			{
				border: 1px dashed #000066;
			}
			.tdbg{
				background:#EEEEEE;
				line-height: 120%;
				font: normal 14px "TArial", "Helvetica", "sans-serif";
			}
			.topbg
			{
				background:#6699cc;
				color: #FFFFFF;
				font: normal 14px "TArial", "Helvetica", "sans-serif";
				text-align: center;
			
			}
			.bgcolor {
				background-color: #BFC1AE;
			}
	</style>
	<p>&nbsp; </p>
	<table width="300" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
		  	<tr align="center"> 
				<td height=25 colspan=2 class="topbg"><strong>系统暂时关闭:</strong></td>
			</tr> 
		    <tr> 
		    <td class="tdbg">
		    <%
		    If Application(cache_name_user&"_systemnote")<>"" Then
		    	Response.Write Application(cache_name_user&"_systemnote")
			Else
				Response.Write "请稍后访问，谢谢。"
			End If
		    %>
		    </td>
		  	</tr>
	 </table>
		<%	
				Response.End
			End If
		End If
	End If
End Sub
%>
