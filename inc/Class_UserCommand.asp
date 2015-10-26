<%@ LANGUAGE = VBScript CodePage = 936%>
<!--#include file="../conn.asp"-->
<!--#include  file="class_sys.asp"-->
<!--#include  file="Inc_Calendar.asp"-->
<!--#include  file="Inc_ubb.asp"-->
<%
Dim oBlog
Set oBlog = New class_sys
oBlog.start

'用户全功能类模块
'本模块目前尽可能与其他模块独立
Class Class_UserCommand
	Public Action
	Public ID
	Public rst			
	Public Title
	Public ErrMsg
	Public mUserSkinLog,mYear,mMonth,mDay
	Private mUserName,mUserId,mUserPath,mUserNickName,mFileName,mUserFolder,mBlogName,mUserPhotoRow,mUsersublist,mUserCmdpath,mUserLogPath
	Private MaxPerPage,Total,strLogN,strUrl
	Private Sql,SqlStart,SqlPart,SqlEnd,rstSubject,strErrMsg
	
	Private Sub Class_Initialize()
		userid=Request("uid")
		'MaxPerPage=5	
	End Sub
		
	Private Sub Class_Terminate()
	    
	End Sub
	
	Public Property Let userid(ByVal Values)
        Dim rstmp, strSql
        mUserid = CLng(Values)
        SqlStart = "Select  * From oblog_log Where userid="& mUserId & " "
		'SqlEnd="  And ishide=0 and passcheck=1 and isdraft=0 and blog_password=0 Order by istop,addtime Desc"
		SqlEnd=" and passcheck=1 and isdraft=0 Order by istop,addtime Desc"
		Action=LCase(Request("do"))		
		Id=Request("Id")
		Call GetUserInfo()
		mFileName=mUserCmdpath&"MyHomestay."&f_ext&"?uid="&mUserid&"&do="
	End Property
	
	Public Function Process()
		Dim strReturn,strMonth,strDay 
		Id=CheckInt(Id)
		strMonth=Request("month")
		strDay=Request("day")
		'Response.Write "动作2：" & Action & "<BR/>" & vbCrlf
		'Response.Write "编号2：" & Id & "<BR/>" & vbCrlf
		Select Case Action
			Case "index"	
				SqlPart=" "	
				Sql=SqlStart &	SqlEnd
				mFileName = mFileName & "index"		
				strReturn = ShowList(Sql,"篇文章","0")
			Case "blogs"
				If Id="" OR Id=0 Then				
					SqlPart=" And logType=0" 
					mFileName = mFileName & "blogs"	
				Else
					SqlPart=" And logType=0 And Subjectid=" & Id
					mFileName = mFileName & "blogs&id=" & Id
				End If
				Sql=SqlStart & SqlPart & SqlEnd
				strReturn = ShowList(Sql,"篇文章","0")
			Case "photos"
				If Id="" OR Id=0 Then				
					SqlPart=" And logType=1" 
					mFileName = mFileName & "photos"
				Else
					SqlPart=" And logType=1 And Subjectid=" & Id
					mFileName = mFileName & "photos&id=" & Id
				End If	
				Sql=SqlStart & SqlPart & SqlEnd
				strReturn = ShowList(Sql,"篇文章","0")		
			Case "month"
				mFileName = mFileName & "month&month=" & strMonth
				strDay=Left(strMonth,4) & "-" & Right(strMonth,2) & "-01" 
				mYear=Left(strMonth,4) 
				mMonth=Right(strMonth,2)			
				If Not IsDate(strDay) Then 
					ErrMsg = "<center>错误的日期数据，应为YYYYMMDD格式，如：20050801</center>"
					Exit Function
				End If
				if is_sqldata=0 then				
					SqlPart=" And Datediff('m',Addtime,'" & strDay &"')=0"
				else
					SqlPart=" And Datediff(m,Addtime,'" & strDay &"')=0"
				end if
				Sql=SqlStart & SqlPart & SqlEnd
				strReturn = ShowList(Sql,"篇文章","0")	
			Case "day"
				mFileName = mFileName & "day&day=" & strDay
				mYear=Year(strDay)
				mMonth=Month(strDay)
				If Not IsDate(strDay) Then 
					strReturn = "<center>错误的日期格式，应为YYYYMMDD格式，如：20050801</center>"
					Exit Function
				End If
				if is_sqldata=0 then
					SqlPart=" And Datediff('d',Addtime,'" & strDay &"')=0"
				else
					SqlPart=" And Datediff(d,Addtime,'" & strDay &"')=0"
				end if
				Sql=SqlStart & SqlPart & SqlEnd
				strReturn = ShowList(Sql,"篇文章","0")	
			Case "message"
				Sql="select * from oblog_message where userid=" & mUserId & " order by messageid desc"
				mFileName = mFileName & "message"	
				strReturn = ShowList(Sql,"个留言","1")	
			Case "comment"
			Case "tag_blogs" '此处将文章与相册合并显示
				mFileName = mFileName & "tag_blogs&id=" & Id
				Sql="Select a.userid,b.* From "
				Sql=Sql & " (Select logid,userid From oblog_usertags Where userid=" & mUserId  & " and tagid=" & id &") a ,"
				'Sql=Sql & " (Select * From oblog_log where userid=" & mUserId & " And logType=0) b Where a.logid=b.logid "	
				Sql=Sql & " (Select * From oblog_log where userid=" & mUserId & ") b Where a.logid=b.logid "
				Sql=Sql & " order By b.addtime Desc"	
				strReturn = ShowList(Sql,"篇文章","0")	
			Case "tag_photos"
				mFileName = mFileName & "tag_photos&id=" & Id
				Sql="Select a.userid,b.* From "
				Sql=Sql & " (Select logid,userid From oblog_usertags Where userid=" & mUserId  & " and tagid=" & id &") a ,"
				Sql=Sql & " (Select * From oblog_log where userid=" & mUserId & " And logType=1) b Where a.logid=b.logid "	
				Sql=Sql & " order By b.addtime Desc"	
				strReturn = ShowList(Sql,"篇文章","0")		
			Case "tags"
				strReturn = GetUserTags()
			Case "show"
				strReturn = ShowOneBlog(Id,0)
			Case "album"
				mFileName = mFileName & "album"
				if id>0 then 
					sql="select file_path,logid from oblog_upfile where isphoto=1 and isPower=0 and userid="&mUserId&" and userClassId="&id&"  order by fileid desc"
				else
					sql="select file_path,logid from oblog_upfile where isphoto=1 and isPower=0 and userid="&mUserId&"  order by fileid desc"
				end if
				strReturn = ShowList(sql,"个图片","2")
			Case Else	
				SqlPart=" "	
				Sql=SqlStart &	SqlEnd
				strReturn = ShowList(Sql,"篇文章","0")					
		End Select
		'strReturn ="oblog_usercontent.innerHTML='" & Replace(Replace(strReturn,"'","\'"),"""","\""") & "';" & VbCrlf & "calendar.innerHTML='" & Replace(Replace(Calendar(2000,1,1),"'","\'"),"""","\""") & "';"
		'strReturn ="oblog_usercontent.innerHTML='" & Replace(strReturn,"'","\'") & "';" & VbCrlf 
		'strReturn= strReturn & "calendar.innerHTML='" & Replace(Calendar(2000,1,1),"'","\'") & "';"
		'strReturn =Replace(strReturn,"""","")
		strReturn=oblog.htm2js_div(filtskinpath(strReturn),"oblog_usercontent")
		Process=strReturn
		'Process="document.write('" & strReturn & "');"
	End Function		
	
	Public Function CreateCalendar()
		Dim strReturn
		If mYear="" Then 
			mYear=Year(Date)
			mMonth=Month(Date)
		End If
		strReturn=oblog.htm2js_div(Calendar(mYear,mMonth,mUserId),"calendar")
		CreateCalendar=strReturn
	End Function
	
	Private Function ShowUserBlogs(rst)
		Dim strBlogs
		Do While Not rst.Eof
			strBlogs= strBlogs & GetOneBlogInfo(rst,"")	& "<BR/>"
			rst.Movenext
		Loop	
		'进行统一处理，不必每篇处理
		strLogMore=Replace(strLogMore,"$show_blogtag$","")
		strLogMore=Replace(strLogMore,"$show_blogzhai$","")		
		strLogMore=Replace(strLogMore,"$show_blogtag","")		
		strLogMore=filt_inc(strLogMore)
		strLogMore=strLogMore & "<script src="""&BlogDir&"count.asp?action=logs&id="&strLogN&"""></script>"			
		'strLogMore=Replace(user_skin_main,"$show_log$",strLogMore)	
		ShowUserBlogs= strLogMore		
		strLogMore=""
		'分页横条，每次只取符合条件的MaxPerPage条，不全部取出
	End Function

	Private Function ShowOneBlog(BlogId,isPower)
		Set rst=oblog.Execute("Select * From oBlog_log Where logid=" & BlogId)
		ShowOneBlog=GetOneBlogInfo(rst,"1")
	End FUnction

	'获取一篇文章的所有内容
	'注意摘要/内容以及尾部标签的处理
	Public Function GetOneBlogInfo(byref rst,byval strMode)
		Dim strTopic,strEmot,strAddtime,strLogtext,strAuthor,strLogInfo,strMore
		Dim strOneLog,strTopictxt,strLogMore,show,rssubject,strTmp,xmlstr,rstmp,strart,i
		'表情
		If rst("face")="0" Then strEmot="" Else	strEmot="<img src="&blogurl&"images/face/" & rst("face") &".gIf />"
		'作者
		If mUserNickName=""  Then 
			strAuthor=mUserName
		Else
			strAuthor=mUserNickName
		End If
		If rst("authorid")<>mUserId Then 
			If Not IsNull(rst("author")) Then
				strAuthor=rst("author")
			End If
		End If
		strAddtime=Year(rst("addtime"))&"-"&Month(rst("addtime"))&"-"&Day(rst("addtime")) &"&nbsp;&nbsp;&nbsp;"&oblog.WeekDateE(rst("addtime"))&""
		strTopic=strEmot
		If rst("istop")=1 Then strTopic="[置顶]"
		If rst("subjectid")>0 Then
			rstSubject.Filter="subjectid=" & rst("subjectid")
			If Not rstSubject.Eof Then
				strTopic=strTopic & "<a href="""& BlogDir & UserPath &"/MyHomestay."&f_ext&"?do=subject&id="">["&oblog.filt_html(rssubject(1))&"]</a>"
			End If
		End If
		strTopictxt="<a href="""& BlogDir & rst("logfile")& """>" & oblog.filt_html(rst("topic")) & "</a>"
		If rst("isbest")=1 Then strTopictxt = strTopictxt & "　<img src=../../images/title03.gif >"
		strTopic = strTopic & strTopictxt
		If rst("istop")=1 Then strTopictxt = "[置顶]" & strTopictxt
		strLogInfo = strAuthor & " 发布于 " & strAddtime
		strMore = "<a href="""& BlogDir & rst("logfile")&""">阅读全文<span id=""ob_logr" & rst("logid") & """></span></a>"
		strMore = strMore&" | "&"<a href=""" & BlogDir & rst("logfile")&"#comment"">回复<span id=""ob_logc" & rst("logid") & """></span></a>"
		strMore = strMore&" | "&"<a href=""../../showtb.asp?id=" & rst("logid") & """ target=""_blank"">引用通告<span id=""ob_logt" & rst("logid") & """></span></a>"
		'摘要
		'If Not IsNull(rst("Abstract")) Then	
		'	strLogtext=rst("Abstract")
		'Else
			strLogtext=oblog.nohtml(rst("logtext"))&"111"'WL
		'End If
		'用来进行计数累计
		strLogN=strLogN&"$"&rst("logid")
		
		'处理内容模板数据
		strOneLog = Replace(mUserSkinLog,"$show_topic$",strTopic)
		strOneLog = Replace(strOneLog,"$show_loginfo$",strLogInfo)
		strOneLog = Replace(strOneLog,"$show_logtext$",strLogtext)
		strOneLog = Replace(strOneLog,"$show_more$",strMore)
		strOneLog = Replace(strOneLog,"$show_emot$",strEmot)
		'strOneLog = Replace(strOneLog,"$show_author$",strAuthor)
		strOneLog = Replace(strOneLog,"$show_addtime$",strAddtime)
		strOneLog = Replace(strOneLog,"$show_topictxt$",strTopictxt)
		strLogMore=strLogMore&strOneLog	
		If strMode="1" Then
			strLogMore=Replace(strLogMore,"$show_blogtag$","")
			strLogMore=Replace(strLogMore,"$show_blogzhai$","")		
			strLogMore=Replace(strLogMore,"$show_blogtag","")		
			'strLogMore=filt_inc(strLogMore)
			strLogMore=strLogMore & "<script src="""&BlogDir&"count.asp?action=logs&id="&strLogN&"""></script>"	
		End If
		GetOneBlogInfo = strLogMore
	End Function

	'用户TAG，不进行分页(Cloud),根据标签查询到的内容不区分文章还是相册
	Private Function GetUserTags()
		Dim sContent,sSql,rst,iFont,iFontSize
	 	sSql = "Select a.TagId,a.Name,b.TagNum From oblog_tags a,"
		sSql = sSql & "(Select Count(*) as TagNum,TagId From oblog_UserTags Where userid=" & mUserId & " Group By TagId ) b Where "
		sSql = sSql & "a.tagid=b.tagid "
	 	'Response.Write sSql
	 	Set rst=conn.Execute(sSql)
	 	If rst.Eof Then
	 		sContent=""
		Else
			Do While Not rst.Eof
				'基数为10
				iFont=rst("TagNum") Mod 10
				If iFont=0 Then iFontSize=9
				If iFont>-1 And iFont<40 Then iFontSize=12 + iFont
				If iFont >40 Then iFontSize=42
				sContent= sContent & "<li><span><a href="""&mUserCmdpath&"MyHomestay."&f_ext&"?uid="&mUserid&"&do=tag_blogs&id=" & rst("tagID") & """><font style=""font-size:"& iFontSize &"px;"">" & rst("Name")& "</font></a></span><br />" 
				sContent= sContent & "<a href="&blogurl&"tags.asp?tagid=" & rst("tagID") &" target=_blank><img src="&blogurl&"images/icon_blogs.gif border=0 title='本站使用过该标签的文章'/></a>"
				sContent= sContent & "<a href="&blogurl&"tags.asp?t=user&tagid=" & rst("tagID") &" target=_blank><img src="&blogurl&"images/icon_users.gif border=0 title='本站使用过该标签的用户'/></a></li>" 
				rst.Movenext
			Loop
		End If
		rst.Close
		Set rst=Nothing
		GetUserTags="<div id=""ob_usertags""><ul>"&sContent&"</ul></div>"
		sContent=""
	End Function

	Private Function ShowList(strSql,strUnit,strMode)
		Dim strReturn
		if action="photos" or action="album" then strReturn="<div id=""albumtop""><ul>"&GetUserClasses(action)&"<ul></div>"
		If Request("page")<>"" Then
	    	CurrentPage=CLNG(Request("page"))
		Else
			CurrentPage=1
		End If
		If Not IsObject(conn) Then link_database
		Set rst=Server.CreateObject("Adodb.RecordSet")
		'Response.Write strSql
		'response.End()
		rst.Open strSql,Conn,1,1
		'Response.Write "符合条件的纪录数目为:" & rst.RecordCount
	  	If rst.Eof  Then
			strReturn=strReturn & "<ul>无记录</ul>"
			ShowList = strReturn
			rst.Close
			Set rst=Nothing
			Exit Function
		End If
		Total=rst.RecordCount
		'strReturn=strReturn & "共调用" & Total & strUnit & "<br>"
		If CurrentPage<1 Then
	     	CurrentPage=1
	    End If
	    If (CurrentPage-1)*MaxPerPage>Total Then
			If (Total mod MaxPerPage)=0 Then
		     	CurrentPage= Total \ MaxPerPage
			Else
			    CurrentPage= Total \ MaxPerPage + 1
		   	End If
	    End If
		If CurrentPage=1 Then
			Select Case strMode
					Case "0"
						strReturn = strReturn&ShowOnePage(rst)
						strReturn=strReturn & oblog.showpage(mFileName,Total,MaxPerPage,false,true,strUnit)
					Case "1"
						strReturn = ShowMessages(rst)
						strReturn="<h1>留言板首页(<a href='"&blogdir&mUserPath&"/message."&f_ext&"#cmt'>签写留言</a>)</h1>" & vbCrLf & strReturn & oblog.showpage(mFileName,Total,MaxPerPage,false,true,strUnit)
					Case "2"
						strReturn = strReturn&getPhotolist(rst)
						strReturn=strReturn & oblog.showpage(mFileName,Total,MaxPerPage,false,true,strUnit)
			End Select	
	   	Else
	   		If (CurrentPage-1) * MaxPerPage < Total Then
	        	rst.move  (CurrentPage-1) * MaxPerPage
	         	'Dim bookmark
	           	'bookmark=rst.bookmark
	            Select Case strMode
					Case "0"
						strReturn = ShowOnePage(rst)
						strReturn=strReturn & oblog.showpage(mFileName,Total,MaxPerPage,false,true,strUnit)
					Case "1"
						strReturn = ShowMessages(rst)
						strReturn="<h1>留言板首页(<a href='"&blogdir&mUserPath&"/message."&f_ext&"#cmt'>签写留言</a>)</h1>" & vbCrLf & strReturn & oblog.showpage(mFileName,Total,MaxPerPage,false,true,strUnit)
					Case "2"
						strReturn = strReturn&getPhotolist(rst)
						strReturn=strReturn & oblog.showpage(mFileName,Total,MaxPerPage,false,true,strUnit) 
				End Select	
	        Else
		        CurrentPage=1
	           	Select Case strMode
					Case "0"
						strReturn = ShowOnePage(rst)
						strReturn=strReturn & oblog.showpage(mFileName,Total,MaxPerPage,false,true,strUnit)
					Case "1"
						strReturn = ShowMessages(rst)
						strReturn="<h1>留言板首页(<a href='"&blogdir&mUserPath&"/message."&f_ext&"#cmt'>签写留言</a>)</h1>" & vbCrLf & strReturn & oblog.showpage(mFileName,Total,MaxPerPage,false,true,strUnit)
					Case "2"
						strReturn = strReturn&getPhotolist(rst)
						strReturn=strReturn & oblog.showpage(mFileName,Total,MaxPerPage,false,true,strUnit) 
				End Select	           	
		    End If
		End If
		rst.Close
		Set rst=Nothing
		ShowList=strReturn
	End Function
	
	Private Function ShowOnePage(rst)
		Dim strBody,strContent,strTmp,rssubject,i,substr
		Dim strTopic,strLoginfo,strLogtext,strMore,strEmot,strAuthor,strAddtime,strTopictxt
		Dim show_log_indexImage
		
		Set rssubject = oblog.execute("select subjectid,subjectname from oblog_subject where userid="&mUserid)
        While Not rssubject.EOF
            substr = substr & rssubject(0) & "!!??((" & rssubject(1) & "##))=="
            rssubject.movenext
        Wend
		Do While Not rst.EOF
			if mUsersublist=1 and id>0 then '列表显示
				strBody="<li><a href="&mUserLogpath&rst("logfile")&" >"&oblog.filt_html(rst("topic"))&"</a>　"&oblog.filt_html(rst("author"))&" <span>"&rst("addtime")&"</span></li>"&vbcrlf				
			else
				If rst("face") = "0" Then 
					strEmot = "" 
				Else 
					strEmot = "<img src="&blogurl&"images/face/" & rst("face") & ".gif />"
				End If
				If mUserNickName = "" Or IsNull(mUserNickName) Then
					strAuthor = mUserName
				Else
					strAuthor = mUserNickName
				End If
	
				If rst("authorid") <> mUserId Then strAuthor = rst("author")
				'strAddtime = rst("addtime")
				strAddtime=Year(rst("addtime"))&"-"&Month(rst("addtime"))&"-"&Day(rst("addtime")) &"&nbsp;&nbsp;&nbsp;"&oblog.WeekDateE(rst("addtime"))&""'WL;
				strTopic = strEmot
				If rst("subjectid") > 0 Then
					strTopic = strTopic & "<a href=""" & mUserCmdpath & "MyHomestay."&f_ext&"?uid="&mUserid&"&do=blogs&id=" & rst("subjectid") & """>[" & oblog.filt_html(getsubname(substr,rst("subjectid"))) & "]</a>"
				End If
				strTopictxt = "<a target=_blank href=""" & mUserLogpath&rst("logfile") & """>" & oblog.filt_html(rst("topic")) & "</a>"
				
				If rst("logPhoto_AddressURL")<>"" And len(rst("logPhoto_AddressURL"))>6 Then'WL;
					show_log_indexImage = "<a target=_blank href='" & rst("logPhoto_AddressURL") &"'>"
					show_log_indexImage = show_log_indexImage & "<img src='"& rst("logPhoto_AddressURL") &"' />"
					show_log_indexImage = show_log_indexImage & "</a>"
				Else'WL;
					show_log_indexImage = ""
				End If'WL;
				
				
				If rst("isbest") = 1 Then strTopictxt = strTopictxt & "　<img src=" & blogurl & "images/title03.gif >"
				strTopic = strTopic & strTopictxt
				If rst("istop") = 1 Then strTopictxt = "[置顶]" & strTopictxt
				'strLoginfo = strAuthor & " 发布于 " & strAddtime
				strLoginfo =  " 发布于 " & strAddtime
				strMore = "<a href=""" & mUserLogpath&rst("logfile") & """>阅读全文("&rst("iis")&")</a>"
				strMore = strMore & " | <a href=""" & mUserLogpath & rst("logfile") & "#cmt"">回复("&rst("commentnum")&")</a>"
				strMore = strMore & " | <a href=""" & blogurl & "showtb.asp?id=" & rst("logid") & """ target=""_blank"">引用通告("&rst("trackbacknum")&")</a>"
				'取得文章摘要内容
				If IsNull(rst("Abstract")) Or Trim(rst("Abstract")) = "" or rst("ishide") = 1 or rst("ispassword") <> "" or rst("passcheck") = 0 Then
					'兼容以前数据
					If rst("ishide") = 1 Then strTmp = "此文章为隐藏文章，仅好友可见，<a href='" & blogurl & "more.asp?id=" & rst("logid") & "'>点击进入验证页面</a>。"
					If rst("ispassword") <> "" Then strTmp = "<form method='post' action='" & blogurl & "more.asp?id=" & rst("logid") & "' target='_blank'>请输入文章访问密码：<input type=""password"" size=""15"" name=""password"" />  <input type=""submit"" value=""提交""></form>"
					If rst("passcheck") = 0 Then strTmp = "此文章需要管理员审核后才可见。"
					If strTmp <> "" Then
						strLogtext = strTmp&"6666666666666666666666"'WL;
					Else
						strLogtext = oblog.nohtml(rst("logtext"))&"333"'WL
						strLogtext = trimlog(strLogtext, rst("showword"))
						'If Left(strLogtext, 7) = "#isubb#" Then
							'strLogtext = UBBCode(strLogtext, 1)
							'strLogtext = Replace(strLogtext, Chr(10), "<br /> ")
							
						'End If
						strLogtext = Replace(strLogtext, "#isubb#", "")
						strLogtext = filtimg(strLogtext)
						If oblog.setup(29, 0) = 1 Then strLogtext = profilthtm(strLogtext)
					End If
				 Else
					'''strLogtext = oblog.nohtml(rst("Abstract"))&"KKKKKKKKKKKKKKKKKKKKKKKKK"'WL''Importent...WL
					strLogtext = Left(oblog.nohtml(rst("logtext")),200)'&"KKKKKKKKKKKKKKKKKKKKKKKKK"'WL''Importent...WL
				 End If
				 strLogtext=oblog.filt_badword(UBBCode(strLogtext,1))
				 '当使用相对路径时，替换为绝对路径
				 'if is_relativepath=1 then
					'	strLogtext=filtskinpath(strLogtext)
				 'end if
				 strlogn = strlogn & "$" & rst("logid")
				 strBody = Replace(mUserSkinLog, "$show_topic$", strTopic)
				 strBody = Replace(strBody, "$show_loginfo$", strLoginfo)
				 strBody = Replace(strBody, "$show_logtext$", strLogtext)
				 strBody = Replace(strBody, "$show_more$", strMore)
				 strBody = Replace(strBody, "$show_emot$", strEmot)
				 strBody = Replace(strBody, "$show_author$", strAuthor)
				 strBody = Replace(strBody, "$show_addtime$", strAddtime)
				 strBody = Replace(strBody, "$show_topictxt$", strTopictxt)	
				 
				 strBody = Replace(strBody, "$show_log_indexImage$", show_log_indexImage)'WL;
				          
				 strBody = Replace(strBody, "$show_blogzhai$", "")	
				 strBody = Replace(strBody, "$show_blogtag$", "")	
				 'show_logmore = show_logmore & strBody
			 end if
	         strContent = strContent & VBCRLF & strBody	         
	         rst.movenext
	         i=i+1
			 if i>=MaxPerPage then exit do
	      Loop
		  set rssubject=nothing
	      ShowOnePage=strContent
		  if mUsersublist=1 and id>0 then
		  	ShowOnePage="<div id=""subject_index""><ul>"&oblog.filt_html(getsubname(substr,id))&ShowOnePage&"</ul></div>"
		  end if
	End Function	
	
	Public Function ShowMessages(rst)
        Dim strtopic, stremot, straddtime, strlogtext, strauthor, strloginfo, strmore, strMessage, strtopictxt, strContent
        Dim homepage_str, user_filepath,i
        If Not rst.EOF Then
            Do While Not rst.EOF
                If IsNull(rst("homepage")) Then
                    homepage_str = "个人主页"
                Else
                    If Trim(Replace(rst("homepage"), "http://", "")) = "" Then
                        homepage_str = "个人主页"
                    Else
                        homepage_str = "<a href=""" & oblog.filt_html(rst("homepage")) & """ target=""_blank"">个人主页</a>"
                    End If
                End If
                strtopic = oblog.filt_html(rst("messagetopic")) & "<a name='" & rst("messageid") & "'></a>"
                If rst("isguest") = 1 Then
                    strauthor = oblog.filt_html(rst("message_user")) & "(游客)"
                Else
                    strauthor = oblog.filt_html(rst("message_user"))
                End If
                straddtime = rst("addtime")
                strtopictxt = strtopic
                strloginfo = strauthor & "发布留言于" & straddtime
                strlogtext = oblog.FilterUbbFlash(filtscript(rst("message")))
                strmore = homepage_str & " | <a href='"&blogurl&"user_messages.asp?action=modify&re=true&id=" & rst("messageid") & "'>回复</a>"
                strmore = strmore & " | <a href=""" & blogurl & "user_messages.asp?action=del&id=" & rst("messageid") & """  target=""_blank"">删除</a>"
                strMessage = Replace(mUserSkinLog, "$show_topic$", strtopic)
                strMessage = Replace(strMessage, "$show_loginfo$", strloginfo)
                strMessage = Replace(strMessage, "$show_logtext$", strlogtext)
                strMessage = Replace(strMessage, "$show_more$", strmore)
                strMessage = Replace(strMessage, "$show_emot$", "")
                strMessage = Replace(strMessage, "$show_author$", strauthor)
                strMessage = Replace(strMessage, "$show_addtime$", straddtime)
                strMessage = Replace(strMessage, "$show_topictxt$", strtopictxt)
                strMessage = Replace(strMessage, "$show_blogtag$", "")
                strMessage = Replace(strMessage, "$show_blogzhai$", "")
                strContent = strContent & strMessage
                rst.movenext
                i=i+1
			 	If i>=MaxPerPage Then Exit Do
            Loop
        Else
            strContent = "暂无留言"
        End If
        ShowMessages=strContent  
  
    End Function
	
	'获取用户信息
	Private Function GetUserInfo()
		Dim rst,rst1,ustr
		Set rst=oBlog.Execute("Select user_folder,user_dir,BlogName,username," &_
			"nickname,user_skin_showlog,user_showlog_num,blog_password,user_photorow_num,user_info From oBlog_User Where UserId=" & mUserId)
		If rst.Eof Then
			Set rst = Nothing
			Response.Write "错误的用户编号"
			Response.End
		Else
			'判断是否整站加密
			if (rst("blog_password")<>"" or isnull(rst("blog_password"))=false) and Request.Cookies(cookies_name)("blogpw")<>rst("blog_password") then
				set rst=nothing
				response.Write "window.location='"&blogurl&"chkblogpassword.asp?userid="&mUserId&"';"
				response.End()
			end if		
			mUserFolder=rst("user_folder")
			mUserPath=rst("user_dir")&"/"&rst("user_folder")
			mBlogName=rst("blogname")
			mUserName=rst("username")		
			mUserNickName=rst("nickname")
			MaxPerPage=rst("user_showlog_num")
			mUserPhotoRow=rst("user_photorow_num")
			ustr=rst("user_info")
			if ustr="" or isnull(ustr) then
				mUsersublist=0
			else
				ustr=split(ustr,"$")
				if ustr(0)<>"" then mUsersublist=cint(ustr(0)) else mUsersublist=0
			end if
			if mUsersublist=1 and id>0 then MaxPerPage=40 '列表模式调用50条
			if mUserPhotoRow<=0 or isnull(mUserPhotoRow) then mUserPhotoRow=1
			If IsNull(rst("user_skin_showlog")) OR rst("user_skin_showlog")="" Then
				Set rst1 = oBlog.Execute("select skinshowlog from oBlog_userskin where isdefault=1")
            	If Not rst1.EOF Then
                	mUserSkinLog = rst1("skinshowlog")
                	Set rst1 = Nothing
            	Else
	                Set rst1 = Nothing
	                Set rs = Nothing
	                Response.Write ("模版错误")
	                Response.End
            	End If
			Else
				mUserSkinLog=rst("user_skin_showlog")
			End If
			if true_domain=1 then 
				mUserCmdpath="/" 
				mUserLogpath=""
			else 
				mUserCmdpath=blogdir&mUserPath&"/"
				mUserLogpath=blogdir
			end if
			'mUserSkinLog=filtskinpath(mUserSkinLog)
		End If		
		Set rst=Nothing
	End Function
	
	Function getPhotolist(rsPhoto)
	Dim i,bstr,n,fso,sReturn
	Dim title,imgsrc
	Set fso = server.CreateObject("Scripting.FileSystemObject")	
	sReturn="<table width='100%'  align='center' cellpadding='0' cellspacing='1'><tbody>"& vbcrlf
	Do While not rsPhoto.eof
		sReturn=sReturn&"<tr>"& vbcrlf
		For n=1 to mUserPhotoRow
			if rsPhoto.eof then
				sReturn=sReturn&"<td width='25%'></td>"& vbcrlf
			Else
				title="<BR/><a href=" & blogurl & "more.asp?id=" & rsPhoto(1) &" target=_blank >阅读图片介绍</a><BR/><BR/>"
				imgsrc=blogurl & rsPhoto(0)
				'imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
				'imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
				'if  not fso.FileExists(Server.MapPath(imgsrc)) then
					'imgsrc=blogurl&rsPhoto(0)
				'End if
				sReturn=sReturn&"<td align='center'> <a href='"& blogurl & rsPhoto(0)&"' target='_blank' title='点击查看原图'><img src='"&imgsrc&"' height='100' width='130' border='0' /></a><br />"&title&"</td>"& vbcrlf
				i=i+1
				if not rsPhoto.eof then rsPhoto.movenext
			End if
		Next
		sReturn=sReturn&"</tr>"& vbcrlf
		if i>=MaxPerPage then exit do	
	loop		
	sReturn=sReturn&"</tbody></table>"	& VBCRLF
	Set fso=nothing
	getPhotolist=sReturn
End Function
	
'获取用户分类
Function GetUserClasses(typestr)
	Dim rst,sReturn,strPlayerUrl
	strPlayerUrl= blogurl & "PhotoPlayer.asp?userid="&muserid
	Set rst=conn.Execute("Select * From oblog_subject Where subjecttype=1 and userid="&mUserid&" order by ordernum")
	If rst.Eof Then
		sReturn=""
	Else
		Do While Not rst.Eof
			sReturn=sReturn&"<option value="&rst("subjectid")&">" & rst("subjectname") & "</option>" & VBCRLF
			rst.Movenext
		Loop
		sReturn = "<option value="""">请选择图片分类</option><option value='0'>所有分类</option>" & VBCRLF & sReturn
		sReturn="<select name=classid onchange=""javascript:window.location='"&mUserCmdpath&"MyHomestay."&f_ext&"?uid="&muserid&"&do="&typestr&"&id='+this.options[this.selectedIndex].value;"">" & VBCRLF & sReturn & "</select>"
	End If
	rst.Close
	Set rst=Nothing
	sReturn=sReturn&"　<a href=""#"" onclick=""window.open('"&strPlayerUrl&"','_photo','height=500, width=480, top=100, left=400, toolbar=no, menubar=no, scrollbars=no, resizable=yes,status=no')"">启用自动播放</a>" & VBCRLF
	if typestr="album" then
		sReturn=sReturn&"  <a href='"&mUserCmdpath&"MyHomestay."&f_ext&"?uid="&mUserid&"&do=photos'>文章方式浏览</a>"
	else
		sReturn=sReturn&"  <a href='"&mUserCmdpath&"MyHomestay."&f_ext&"?uid="&mUserid&"&do=album'>相册方式浏览</a>"
	end if 
	GetUserClasses = sReturn	
End Function
	
End Class	
%>