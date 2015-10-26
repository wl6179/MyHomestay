<%
'*********************************************************
'File: 			Inc_Tags.asp
'Description:	Tag Functions For oBlog3.1
'Author: 		王旭 
'Copyright:		http://www.myhomestay.com.cn
'LastUpdate:	20050921
'*********************************************************
Sub Tags_UserAdd(byval sTags,byval sUserId,byval slogId)
	If sTags="" Then Exit Sub
	Dim i,j,iCount,aTags,TagId,sTag,sTagsId
	Dim rst
	Set rst=Server.CreateObject("ADODB.RECORDSET")
	aTags=Split(sTags,P_TAGS_SPLIT)
	j=0
	For i=0 To UBound(aTags)
		sTag = aTags(i)
		sTag=EncodeJP(ProtectSql(sTag))
		'Must >=2 And Not bad One
		If Len(sTag)>1 And Tags_CheckTag(sTag)=1 Then			
			'For System		
			rst.Open "Select * From oblog_Tags Where Name='" & Trim(aTags(i)) & "'",conn,1,3
			If rst.RecordCount=0 Then
				rst.AddNew 
				rst("Name")=sTag
				rst("iNum")=1			
				rst("LastUpdate")=ServerDate(Now)
				rst.Update
				rst.close
				rst.Open "Select * From oblog_Tags Where Name='" & Trim(aTags(i)) & "'",conn,1,3
				if rst.eof then	
					TagId=0
				else
					TagId=rst("TagId")
				end if
			Else
				TagId=rst("TagId")
				rst("iNum")= rst("iNum")+1
				rst("LastUpdate")=ServerDate(Now)
				rst.Update			
			End If
			rst.close
			'For Users		
			rst.Open "Select * From oblog_UserTags Where UserId=" & sUserId &" And logid= " & slogId &" And TagId=" & TagId ,conn,1,3
			If rst.RecordCount=0 Then
				rst.AddNew
				rst("userid")=sUserId
				rst("logid")=slogId
				rst("tagid")=TagId
				rst("iNum")=1	
				rst.Update
			End If
			rst.close
			j=j+1
			If j=1 Then 
				sTags=sTag
				sTagsId=TagId
			Else
				sTags=sTags & P_TAGS_SPLIT & sTag 
				sTagsId= sTagsId & P_TAGS_SPLIT & TagId
			End If
		End If
	Next
	'Update two fields
	Call conn.Execute("Update oblog_log Set logtags='" & sTags &"',logtagsid='" & sTagsId &"' Where logid=" & slogId)
	Set rst=Nothing
End Sub

Sub Tags_UserEdit(byval sTags,byval sUserId,byval slogId)
	'Delete all first
	Call Tags_UserDelete(slogId)
	'Re-Add
	Call Tags_UserAdd(sTags,sUserId,slogId)
End Sub

Sub Tags_UserDelete(byval slogId)	
	Dim i,j,iCount,sTags,TagId
	Dim rst
	Set rst=Server.CreateObject("ADODB.RECORDSET")
	'Get All tags from this blog
	rst.Open "Select tagid From oblog_UserTags Where logid In (" & slogId & ")" ,conn,1,1
	If rst.Eof Then
		rst.Close
		Set rst=Nothing
		Exit Sub
	End If
	Do While Not rst.Eof
		'update  number
		Call conn.Execute("Update oblog_Tags Set iNum=INum-1 Where TagId=" & rst("tagId"))
		rst.Movenext
	Loop
	rst.Close	
	Set rst = Nothing	
	Call conn.Execute("Delete " & delchr & " From oblog_UserTags Where logid In (" & slogId & ")" )		
	Call conn.Execute("Update oblog_log Set logtags='',logtagsid='' Where logid In (" & slogId & ")")	
End Sub

'tag infomation for blog 
Function Tags_ShowForBlog(byval sBlogid,utruepath)
	Dim i,aTags,aTagsId,sContent,sTags,sTagsId,sUserId
	Dim rst
	Set rst=conn.Execute("Select logtags,logtagsid,userid From oblog_log Where logid=" & sBlogid )
	If rst.Eof Then 
		Set rst=Nothing
		Exit Function
	End If	
	If IsNull(rst(0)) Then  Exit Function
	If IsNull(rst(1)) Then  Exit Function
	sTags=Trim(rst(0))
	sTagsId=Trim(rst(1))
	sUserId=rst(2)
	Set rst=Nothing
	If sTags="" OR  sTagsId="" Then
		Tags_ShowForBlog=""
		Exit Function
	End If
	aTags=Split(sTags,P_TAGS_SPLIT)
	aTagsId=Split(sTagsId,P_TAGS_SPLIT)
	For i=0 To UBound(aTags)
		sContent = sContent&"&nbsp;<a href="""& utruepath & "MyHomestay."&f_ext&"?uid=" & sUserId & "&do=tag_blogs&id=" & aTagsId(i) & """>" & aTags(i) & "</a><a href="""& blogurl & P_TAGS_SYSURL & "?tagid=" & aTagsId(i) & """><img src="""&blogurl&"images/icon_blogs.gif"" alt=""查看本站使用过本标签的文章""></a><a href="""& blogurl & P_TAGS_SYSURL & "?t=user&tagid=" & aTagsId(i) & """><img src="""&blogurl&"images/icon_users.gif"" alt=""查看本站使用过本标签的用户""></a>" 
	Next
	Tags_ShowForBlog = sContent
End Function

'user tag list
Function Tags_UserTags(byval sUserId)
	Dim sContent,sSql,rst
	Set rst=Server.CreateObject("ADODB.RECORDSET")
	sSql = "Select top 10 a.TagId,a.Name,b.TagNum From oblog_tags a,"
	sSql = sSql & "(Select Count(*) as TagNum,TagId From oblog_UserTags Where userid=" & sUserId & " Group By TagId ) b Where "
	sSql = sSql & "a.tagid=b.tagid Order By b.TagNum Desc"	
	rst.Open sSql,conn,1,1
	If rst.Eof Then
		sContent=""
	Else
		Do While Not rst.Eof
			sContent=sContent & "<font class=tag1><a href=""cmd.asp?do=tag_blogs&id=" & rst("TagId") & "&uid=" & sUserId &""">" & rst("Name") & "</a></font>(" & rst("TagNum") & ")<BR/>" & VBCRLF
			rst.MoveNext
		Loop
	End If
	rst.Close
	Set rst=Nothing
	Tags_UserTags=sContent
	sContent=""
End Function

'user blog list with tag keyword
Function Tags_TagBlogs(byval sUserId,byval sTagId)
	Dim sContent,sSql,rst
	sSql="Select a.userid,b.* From "
	If sUserId<>"" Then		
		sSql=sSql & " (Select logid,userid From oblog_usertags Where userid=" & sUserId  & " and tagid=" & sTagId &") a ,"
		sSql=sSql & " (Select topic,addtime,logid,author,logfile From oblog_log where userid=" & suserId & ") b Where a.logid=b.logid "		
	Else		
		sSql=sSql & " (Select logid,userid From oblog_usertags Where  tagid=" & sTagId &") a ,"
		sSql=sSql & " (Select topic,addtime,logid,author,logfile From oblog_log ) b Where a.logid=b.logid "		
	End If
	sSql=sSql & " order By b.addtime Desc"	
	Set rst=conn.Execute(sSql)	
	If rst.Eof Then
		sContent=""
	Else
		Do While Not rst.Eof
			sContent=sContent&"<li>&nbsp;&nbsp;<a href=""go.asp?userid="& rst("userid") & """ target=_blank> " & rst("author") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=go.asp?logid="&rst("logid")&" target=""_blank"">"& rst("topic")&"</a><span>&nbsp;&nbsp;&nbsp;&nbsp;"&rst("addtime")&"</span></li>"& VBCRLF
			rst.movenext
		Loop
		If sContent="" Then sContent="<UL>" & VBCRLF & sContent & "</UL>" & VBCRLF 
	End If
	Set rst=Nothing
	Tags_TagBlogs=sContent
	sContent=""
End Function

Function Tags_HotTags()
	Dim sContent,rst,sSql,i
	Set rst=Server.CreateObject("ADODB.RECORDSET")
	sContent="<TABLE id=tab_all cellSpacing=0 cellPadding=0 width=100% border=0>" & VBCRLF
  	sContent= sContent& " <TBODY><TR id=tr_tit> <TD width=30% >" & P_TAGS_DESC & "</TD><TD>文章引用总数</TD><TD>使用人气指数</TD><TD>最后更新时间</TD></TR>" & VBCRLF    
  	sSql="Select a.*,b.UserNum From (Select top 100 tagId,Name,iNum,LastUpdate From oblog_Tags Where iNum>0 And iState=1) a,"
  	sSql=sSql & " (Select Count(*)as  UserNum ,tagid From oblog_UserTags  Group By tagid) b Where a.TagId = b.TagId Order By a.iNum Desc"  	
  	rst.Open sSql,conn,1,1
  	i=0
  	Do While Not rst.Eof
  		If i Mod 2 =0 Then
  			sContent= sContent& " <TR id=tr_nor2>" & VBCRLF
  		Else
  			sContent= sContent& " <TR id=tr_nor1>" & VBCRLF
  		End If
  		sContent= sContent& "<TD align=left>&nbsp;&nbsp;<a href=""" & blogurl & P_TAGS_SYSURL &"?tagid=" & rst("tagid") & """ target=_blank>" & rst("Name") & "</a></TD>"
  		sContent= sContent& "<TD align=left>&nbsp;&nbsp;" & rst("iNum") & "</TD>"
  		sContent= sContent& "<TD align=left>&nbsp;&nbsp;" & rst("UserNum") & "</TD>"
  		sContent= sContent& "<TD align=left>&nbsp;&nbsp;" & rst("LastUpdate") & "</TD></TR>"
  		i=i+1
  		rst.movenext
  	Loop  
  	rst.close
  	Set rst=Nothing
  	sContent= sContent& "</TBODY></TABLE>"
  	Tags_HotTags=sContent
  	sContent=""
 End Function

Function Tags_SearchTag(byval sTagName)
	Dim sSql,rst,sContent,i
	sTagName=EncodeJP(sTagName)
	sSql="Select * From (Select * From oblog_tags Where name Like '%" & sTagName & "%' And iState=1) a Order By a.iNum Desc"
	Set rst=conn.Execute(sSql)	
	If rst.Eof Then
		sContent="没有查询到包含<font color=red><b>" & sTagName & "</b></font>" & P_TAGS_DESC
	Else
		i=1
		Do While Not rst.Eof
			sContent= sContent & i & ":&nbsp;&nbsp;<a href=""" & blogurl & P_TAGS_SYSURL & "?tagid=" & rst("tagid") & """ target=""_blank"">" & Replace(rst("name"),sTagName,"<font color=red>" & sTagName & "</font>") & "</a>(" & rst("iNum") & ")<BR/>" & VBCRLF
			rst.Movenext
		Loop
	End If	
	Set rst=Nothing
	Tags_SearchTag=sContent
	sContent=""
End Function

Function Tags_TagName(byval sTagId,byref o_TagName,byref o_Num,byref o_LastUpdate)
	Dim rst
	Set rst=conn.Execute("Select top 1 * From oblog_Tags Where tagid=" & sTagid & " And iState=1")
	If rst.Eof Then
		Tags_TagName=-1
	Else
		o_TagName=rst("Name")
		o_Num=rst("iNum")
		o_LastUpdate=rst("LastUpdate")
		Tags_TagName=0
	End IF
	Set rst=Nothing
End Function
'Tag Cloud
'1:Cloud;0:List
Function Tags_SystemTags(byval t)
	Dim sContent,sSql,rst,iFont,iFontSize,i
	Dim sSplit
	sSplit="&nbsp;&nbsp;&nbsp;&nbsp;" & VBCRLF
 	sSql="Select top 100 * From oblog_Tags Where iNum>0 "
 	If t=0 Then sSql= sSql & " Order By iNum Desc"
 	Set rst=conn.Execute(sSql)
 	If rst.Eof Then
 		sContent=""
	Else
		Do While Not rst.Eof
			If t=0 Then
				sContent= sContent & "<font class=tag0><a href=""" & blogurl & P_TAGS_SYSURL  & "?tagid=""" & rst("tagID") &""">" & rst("Name")& "(" & rst("iNum") &  ")</a></font>" & sSPlit
			Else
				iFont=rst("iNum") Mod 100
				If iFont=0 Then iFontSize=10
				If iFont>-1 And iFont<40 Then iFontSize=12 + iFont
				If iFont >40 Then iFontSize=42
				sContent= sContent & "<a href=""" & blogurl & P_TAGS_SYSURL  & "?tagid=" & rst("tagID") & """><font style=""font-size:"& iFontSize &"px;line-height:42px"">" & rst("Name")& "</font></a>" & sSPlit
			End If
			i=i+1
			'If i Mod P_TAGS_PerLine = 0 Then
				'sContent = sContent &  "<BR/>"
			'End If			
			rst.Movenext
		Loop
	End If
	rst.Close
	Set rst=Nothing
	Tags_SystemTags=sContent
	sContent=""
End Function

'Forbid Tag For Admin
Function Tags_SystemForbid(byval sTagId)
	conn.Execute("Update oblog_Tags Set iState=0 Where TagId=" & sTagId)
End Function

'Check Bad tags
Function Tags_CheckTag(byval sTag)
	Dim aBadTags,i,lNumber
	'进行过滤
	'aBadTags=Split(oblog.setup(80,0),vbcrlf)
	'Tags_CheckTag =1
	'Bad Tag is shortter than True Tag
	'For i=0 To Ubound(aBadTags)
	'	If Instr(sTag,aBadTags(i)) Then
	'		Tags_CheckTag= 0
	'		Exit Function
	'	End If
	'Next	
	'进行特殊字符的保护
	For i=1 To Len(sTag)
		lNumber=ASC(Mid(sTag,i,1))		
		If lNumber<0 OR (lNumber>=45 And lNumber<=90) OR (lNumber>=97 And lNumber<=122) Then			
		Else
			Tags_CheckTag= 0
			Exit Function
		End If
	Next
	Tags_CheckTag =1
End Function

Function GetUsersByTag(byval sTagId)
	Dim rst,sSql,sContent
	Set rst = Server.CreateObject("Adodb.Recordset") 
	sSql = "Select Top 100 b.userName,b.user_dir,b.user_folder From (Select Userid From oblog_usertags Where Tagid=" & sTagId & " Group By UserId) a,oblog_user b Where a.Userid=b.UserId"
	rst.Open sSql,conn,1,1
	If rst.Eof Then
		sContent="没有符合条件的用户"
		rst.Close
		Set rst = Nothing
	End If
	i=0
	Do While Not rst.Eof
		sContent=sContent & "<a href="& blogurl& rst("user_dir") & "/" & rst("user_folder")&"/index." &f_ext&" target=_blank>" & rst("userName") & "</a><br/>"
		rst.movenext
	Loop
	rst.Close
	Set rst = Nothing
	GetUsersByTag=sContent
End Function

Function GetUserInfo(byval sUserId)
	Dim rst,sUserPath	
	Set rst = Server.CreateObject("Adodb.Recordset") 
	rst.Open "Select * From oblog_User Where Userid=" & sUserId,conn,1,1
	If rst.Eof Then
		GetUserInfo="错误的用户信息"
	Else
		sUserPath= blogurl & rst("user_dir") & "/" & rst("user_folder") & "/index." & f_ext
		sUserPath=Replace(sUserPath,"//","/")
		GetUserInfo="<a href="""  & sUserPath & """ target=_blank>" & rst("blogname") & "</a>"
	End If
	rst.Close
	Set rst = Nothing
End Function

%>