<%
'*********************************************************
'File: 			Inc_Functions.asp
'Description:	���ú���ģ�� For 3.1
'Author: 		���� 
'Copyright:		http://www.myhomestay.com.cn
'LastUpdate:	20050921
'*********************************************************

'	ProIndexLog
'���������
'	strHTML Logtext����
'	iLen  ��ȡ�ĳ��ȣ���HTML����Ϊ������Ҫ��Ϊ�˺���f�ķ�����ʾ
'	iLine Ĭ����ʾ����
Public Function ProIndexLog(strHTML,iLen,iLine) 
	ON ERROR RESUME NEXT
	Dim aHTML,i,j,lP1,lP2	
	Dim objRegExp, strOutput,sOut
	
	If Instr(strHTML,"#��ǰ����ҳ������ʾ#") Then
		strOutput=Split(strHTML,"#��ǰ����ҳ������ʾ#")(0) 
	Else
		strHTML=Left(strHTML,iLen)
		lP1=InStrRev(strHTML,"<")
		lP2=InStrRev(strHTML,">")
		'ȡ500�ı��������ж����һ��<���ֵ�λ�ã������<֮��û�г���>
		'�򲹳�һ�����бպ�
		If lP1>0 Then
			If lP2<lP1 Then
				strHTML=strHTML & ">"
			End If
		End If
		'ȷ�϶���λ��
		strHTML=Replace(strHTML,"<BR>","$breakline$")
		strHTML=Replace(strHTML,"<BR/>","$breakline$")
		strHTML=Replace(strHTML,"<br/>","$breakline$")
		strHTML=Replace(strHTML,"<br />","$breakline$")
		strHTML=Replace(strHTML,"<br>","$breakline$")
		strHTML=Replace(strHTML,"</p>","$breakline$")
		strHTML=Replace(strHTML,"</P>","$breakline$")
		'�޳�<>���
		Set objRegExp = New Regexp'����������ʽ��
		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		objRegExp.Pattern = "<.+?>"
		strOutput = objRegExp.Replace(strHTML, "")
		Set objRegExp = Nothing	
		'��$breakline$�ָȡiLine��
		aHTML=Split(strOutput,"$breakline$")
		j=0
		For i=0 to Ubound(aHTML)
			j=j+1
			strOutput=	Trim(aHTML(i))
			If strOutput<>"" Then
				If j<iLine then
					sOut=sOut & strOutput & "<BR/>" & Vbcrlf
				else			
					'�ڵ�iLine�д������ж�,������д���10��������ʾ��������ʾ����
					'��ֹ�����ƻ���ʾ���
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
		strOutput=Replace(sOut,"��","��<BR/>")		
		If UCase(Right(strOutput,4))="<BR>" Then
			strOutput = Left(strOutput, Len(strOutput)-4)
		ElseIf UCase(Right(strOutput,5))="<BR/>" Then
			strOutput = Left(strOutput, Len(strOutput)-5)
		End If
	End If
	'����֮����л��б�Ǵ˴�д�ں����ⲿ			
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
'���ı���
'10k���±������С��0.01�룬����Ӱ�쵽ִ��Ч��
'Ŀǰ��Ҫ���µ�λ��Ϊ��
'վ��������ĸ�����Ŀ�����ơ�����
'��������ʱ�ı��⡢���ݡ��ؼ���
'��������/����ʱ������
'����ʱ�Թؼ��ֽ��б���
'��ʱ������ע��������
'���������������ʹ��
'------------------------------------------------
Function EncodeJP(byval strContent)

	If strContent="" Then Exit Function
	
	'SQL�汾�����б���
	If IS_SQLDATA=1 Then 
		EncodeJP=strContent
		Exit Function
	End If
	
	strContent=Replace(strContent,"��","&#12460;")
    strContent=Replace(strContent,"��","&#12462;")
    strContent=Replace(strContent,"��","&#12464;")
    strContent=Replace(strContent,"��","&#12450;")
    strContent=Replace(strContent,"��","&#12466;")
    strContent=Replace(strContent,"��","&#12468;")
    strContent=Replace(strContent,"��","&#12470;")
    strContent=Replace(strContent,"��","&#12472;")
    strContent=Replace(strContent,"��","&#12474;")
    strContent=Replace(strContent,"��","&#12476;")
    strContent=Replace(strContent,"��","&#12478;")
    strContent=Replace(strContent,"��","&#12480;")
    strContent=Replace(strContent,"��","&#12482;")
    strContent=Replace(strContent,"��","&#12485;")
    strContent=Replace(strContent,"��","&#12487;")
    strContent=Replace(strContent,"��","&#12489;")
    strContent=Replace(strContent,"��","&#12496;")
    strContent=Replace(strContent,"��","&#12497;")
    strContent=Replace(strContent,"��","&#12499;")
    strContent=Replace(strContent,"��","&#12500;")
    strContent=Replace(strContent,"��","&#12502;")
    strContent=Replace(strContent,"��","&#12502;")
    strContent=Replace(strContent,"��","&#12503;")
    strContent=Replace(strContent,"��","&#12505;")
    strContent=Replace(strContent,"��","&#12506;")
    strContent=Replace(strContent,"��","&#12508;")
    strContent=Replace(strContent,"��","&#12509;")
    strContent=Replace(strContent,"��","&#12532;")

    EncodeJP=strContent
End Function

'------------------------------------------------
'Pause(byval iCount)
'��ͣ����
'��������ת�ƣ�ת�������ɹ����У���ֹ�����ķ�ϵͳ��Դ
'------------------------------------------------
Sub Pause()
	Dim i,lStep,iCount
	iCount=P_BLOG_UPDATEPAUSE
	'��������ִ��ʱ��Ϊ0.03~0.05
	lStep=200000	
	'���Ϊ0���߷���ֵ������
	If  Not IsNumeric(iCount) OR iCount=0 Then Exit Sub
	iCount=CLNG(iCount)
	'Response.Write timer & "<br>"
	'��������3~5��
	If iCount>100 Then iCount=100
	For i=1 To iCount * lStep
	Next
	'Response.Write timer 
End Sub

'------------------------------------------------
'CheckValidEnName(byval strName)�����Ч��Ӣ��IDע���ʺţ�ΪӢ����ĸ�����ֺ��»��ߣ�
'ֻ��������(48~57)+��(65~90)С(97~122)д��ĸ���»���
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
	'���������ȫ��Ϊ���ֵ�ID��wl
	If en_nameisnum=0 Then
		If IsNumeric(strName) Then 	CheckValidEnName=False		
	End If
	Set objReg=Nothing
End Function

'------------------------------------------------
'FilterJS(strHTML)
'���˽ű�
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
'��鲢ת������ֵ
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
'���ڽ��յ�ַ����������ʱSQL��ϱ���
'------------------------------------------------
'��ֹSQLע��
Function ProtectSQL(sSql)
	If ISNull(sSql) Then Exit Function
	sSql=Trim(sSql)
	If sSql="" Then Exit Function 
	sSql=Replace(sSql,Chr(0),"")	
	sSql=Replace(sSql,"'","��")
	sSql=Replace(sSql," ","")
	sSql=Replace(sSql,"%","��")
	sSql=Replace(sSql,"-","��")		
	ProtectSQL=sSql
End Function

'�����û������ĸ�����Ϣ���ˣ����໰����
Function HTMLEncode(fString)
	If Not IsNull(fString) Then
		fString = replace(fString, ">", "&gt;")
		fString = replace(fString, "<", "&lt;")
		fString = Replace(fString, CHR(32), " ")		'&nbsp;
		fString = Replace(fString, CHR(9), " ")			'&nbsp;
		fString = Replace(fString, CHR(34), "&quot;")
		'fString = Replace(fString, CHR(39), "&#39;")	'�����Ź���
		fString = Replace(fString, CHR(13), "")
		fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
		fString = Replace(fString, CHR(10), "<BR> ")
		'fString=ChkBadWords(fString)
		HTMLEncode = fString
	End If
End Function

'------------------------------------------------	
'RemoveHtml(byval strConten
'�Ƴ�HTML���
'��Ҫ�û����浽���ݿ�ǰ�Ĺ���
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
'��������������301�ض�����������Ŀ���ַ
'------------------------------------------------
Sub RedirectBy301(byval strURL)
	Response.Clear
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location",strURL
	Response.End
End Sub

'------------------------------------------------
'ServerDate(byval strDate)
'������ʱ������
'�ظ�/���Լ���������
'����Trackback
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

'������ʹ�ô˷�����include������Ҫ��
Function ReadFileToString(byval oFSO,byval userpath,byval sFile)
'��Ŀ¼���д���
'���ļ��Ǵ���ײ���ʼ��
'On Error Resume Next
	Dim oStream
	'��������inc
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
	'�˴���ʱ�����ж��ļ��Ƿ���ʵ����
	Set oStream=oFSO.OpenTextFile(Server.Mappath(sFile),1,False)
	ReadFileToString = oStream.ReadAll
	Set oStream=Nothing
	'If Err.Number>0 Then ReadFileToString=""
End Function

'��ȡ������IP
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
				<td height=25 colspan=2 class="topbg"><strong>ϵͳ��ʱ�ر�:</strong></td>
			</tr> 
		    <tr> 
		    <td class="tdbg">
		    <%
		    If Application(cache_name_user&"_systemnote")<>"" Then
		    	Response.Write Application(cache_name_user&"_systemnote")
			Else
				Response.Write "���Ժ���ʣ�лл��"
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
