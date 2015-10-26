<%
'*********************************************************
'File: 			Class_Trackback.asp
'Description:	Trackback Class For oBlog3.1
'Author: 		ÍõÐñ
'Copyright:		http://www.myhomestay.com.cn
'LastUpdate:	20050923
'*********************************************************
'Reference
'http://www.sixapart.com/pronet/docs/trackback_spec
'a ping request might look like:
'POST http://www.example.com/trackback/5
'Content-Type: application/x-www-form-urlencoded; charset=utf-8 
'title=Foo+Bar&url=http://www.bar.com/&excerpt=My+Excerpt&blog_name=Foo

Class Class_TrackBack	
	Public LogId
	Public ID
	Public URL
	Public Title
	Public Blog_Name
	Public Excerpt
	Public IP
	Public Agent

	Private Function SendResult(strMsg)
		Dim strXML
		strXML="<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><response><error>%e</error><message>%m</message></response>"

		If strMsg="undiscovered" Then
			strXML=Replace(strXML,"%e","1")
			strXML=Replace(strXML,"%m",strMsg)
		ElseIf strMsg="repetition" Then
			strXML=Replace(strXML,"%e","1")
			strXML=Replace(strXML,"%m",strMsg)
		Elseif strMsg="invalid parameter" Then
			strXML=Replace(strXML,"%m",strMsg)
		Elseif strMsg="none data" Then
			strXML=Replace(strXML,"%e","1")
			strXML=Replace(strXML,"%m",strMsg)
		Else	
			strXML=Replace(strXML,"%e","0")
			strXML=Replace(strXML,"%m",strMsg)
		End If
		Response.ContentType = "text/xml"
		Response.Clear
		Response.Write strXML
		
	End Function


	Public Function Receive()
		Dim UserId
		logId=CheckInt(LogId)		
		IP=GetIP		
		URL=ProtectSQL(URL)
		Title=ProtectSQL(Title)
		Blog_Name=ProtectSQL(Blog_Name)
		Excerpt=ProtectSQL(Excerpt)
		Response.Write logId & "<BR/>"
		Response.Write IP & "<BR/>"
		Response.Write URL & "<BR/>"
		Response.Write Title & "<BR/>"
		Response.Write Blog_Name & "<BR/>"
		Response.Write Excerpt & "<BR/>"
		
		If LogId=0 Then 
			Call SendResult("invalid parameter")
			Receive=False
			Exit Function
		End if
		
		If Len(URL)=0 Then 
			Call SendResult("none data")
			Receive=False
			Exit Function
		End If
		
		If Len(URL)>255 Then 
			Call SendResult("url is long")
			Receive=False:Exit Function
		End If
			
		If Len(Blog_Name)>255 Then 
			Call SendResult("name is long")
			Receive=False
			Exit Function
		End If
		If Len(Blog_Name)=0 Then Blog_Name="Unknow"
		If Len(Excerpt)=0 Then Excerpt=""
		If Len(Excerpt)>255 Then Excerpt=Left(Excerpt,252)&"..."
		If Len(Title)>255 Then Title=Left(Title,252)&"..."
		If Len(Title)=0 Then Title=URL

		Dim rst
		Set rst=conn.Execute("Select * From [oblog_log] Where LogId=" & LogId)
		If rst.Eof Then
			Call SendResult("undiscovered")
			Exit Function
		End If
		Userid=rst("userid")
		rst.Close
		'Set rst=Nothing
		'Response.Write URL & "<br/>"
		'Response.Write ("Select * From [oblog_TrackBack] Where [LogId]=" & LogId & " and tb_url='" & URL & "'")
		'response.End()
		'Ping
		rst.Open "Select * From [oblog_TrackBack] Where [LogId]=" & LogId & " and tb_url='" & URL & "'",conn,1,3
		If Not rst.bof Then
			rst.close
			Call SendResult("repetition")
			Exit Function
		Else
			rst.AddNew
			'rst("userId")=UserId
			rst("logId")=LogId
			rst("topic")=Title
			rst("Blog_Name")=Blog_Name
			rst("Excerpt")=Excerpt
			rst("Url")=URL
			rst("IP")=IP
			rst.Update
		End If
		rst.Close
		Set rst=Nothing
		conn.execute("update [oblog_log] set trackbacknum=trackbacknum+1 where logid="&LogId)		
		Call SendResult("succeed")
		
		Receive=True

	End Function

	Public Function DeleteTrackBack()
		If IsNumeric(ID) Then LogId=CLNG(ID)
		conn.Execute("Delete "& delchar & " From [oblog_TrackBack] Where id =" & ID)
		DeleteTrackBack=True
	End Function	
	
	Private Function Ping(strTarget)

		Dim strSend,objPing
		strSend = "&title=" & Server.URLEncode(Title) & "&url=" & Server.URLEncode(URL) & "&excerpt=" & Server.URLEncode(Excerpt) & "&blog_name=" & Server.URLEncode(Blog_Name)
		
		Set objPing = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objPing.open "POST",strTarget,False
		objPing.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		objPing.send strTarget&strSend
		'Response.Write strTarget&strSend
		Set objPing = Nothing
		Ping=True

	End Function	
	
	Public Function ProcessMultiPing(strTargetUrls)
		Dim aUrls,i,strUrl,rst
		strTargetUrls=Trim(strTargetUrls)
		If strTargetUrls="" Then Exit Function
		aUrls=Split(strTargetUrls,VBCRLF)
		For i=0 To UBound(aUrls)
			strUrl=Lcase(aUrls(i))
			If Left(strUrl,7)="http://" Then
				Call Ping(strUrl)
			End If
			If i+1> P_TRACKBACK_MAX Then Exit Function
		Next
	End Function
End Class
%>