<!--#include file="../inc/Inc_UBB.asp"-->
<!--#include file="../inc/Inc_Tags.asp"-->
<%
'*********************************************************
'�ļ�����: Class_blog.asp
'��������: ���²�������ģ�� 
'��������������,����
'�ٷ���վ: http://www.myhomestay.com.cn
'60������+����֧�֣�http://www.myhomestay.com.cn
'Copyright (C) 2004-2005 MyHomestay.com.cn All rights reserved.
'LastUpdate:    20050921
'*********************************************************

Class Class_Blog
    Public GoUrl, user_skin_main,user_skin_main_Nav, user_skin_showlog, user_userName, user_id, user_nickName, user_showName
    Public user_commentasc, user_path, user_folder,user_showlog_num, user_showlogword_num, BlogName, user_siteinfo,user_truepath
    Public user_Blogpassword, user_domain, user_placard, user_links, user_log_count, user_comment_count
    Public user_message_count, user_shownewlog_num, user_shownewmessage_num, user_shownewcomment_num,log_truepath,user_level,user_logclassid_Nav
    Public rs, objFSO, tf, ispwBlog,showpwblog,showpwlog,Page
    Public m_index,m_log,m_subjectid,m_subjectindex,m_message,m_album,m_info,m_placard,m_links,m_newblog,m_newmessage,m_comment,m_subject,m_subject_l,m_commentsmore

    Private Sub Class_Initialize()
        Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
        showpwblog=false
		showpwlog=false
    End Sub
    
    Private Sub Class_Terminate()
        Set objFSO = Nothing
        Set tf = Nothing
        Set rs = Nothing
    End Sub
    
    Public Property Let userid(ByVal Values)'�������ԣ�
        Dim rstmp, strSql
        userid = CLng(Values)
        strSql = "Select user_dir,user_showlog_num,user_showlogword_num,user_skin_main,user_skin_showlog,"
        strSql = strSql & "BlogName,nickName,userName,siteinfo,Blog_password,"
        strSql = strSql & "comment_isasc,user_domain,user_domainroot,user_placard,user_links,"
        strSql = strSql & "log_count,comment_count,message_count,user_shownewlog_num,user_shownewcomment_num,"
        strSql = strSql & "user_shownewmessage_num,user_folder,user_level"&str_domain&",logclassid_Nav,user_skin_main_Nav"
        strSql = strSql & " From oBlog_user Where userid=" & userid
        Set rs = oblog.Execute(strSql)
        If rs.EOF Then Exit Property
        user_id = userid
        user_path = Trim(rs("user_dir")) & "/" & rs("user_folder")
        user_showlog_num = CLng(rs("user_showlog_num"))
        user_showlogword_num = rs("user_showlogword_num")
        user_skin_main = Trim(rs("user_skin_main"))
		
		user_skin_main_Nav = Trim(rs("user_skin_main_Nav"))
        user_skin_showlog = Trim(rs("user_skin_showlog"))
        BlogName = oblog.filt_html(rs("BlogName"))
        user_nickName = oblog.filt_html(rs("nickName"))
        user_userName = oblog.filt_html(rs(7))
        user_siteinfo = oblog.filt_html(rs(8))
        user_Blogpassword = Trim(rs(9))
        user_commentasc = rs(10)
        user_domain = Trim(rs(11)) & "." & Trim(rs(12))
        user_placard = rs(13)
        user_links = rs(14)
        user_log_count = rs(15)
        user_comment_count = rs(16)
        user_message_count = rs(17)
        user_shownewlog_num = rs(18)
        user_shownewcomment_num = rs(19)
        user_shownewmessage_num = rs(20)
        user_folder=rs("user_folder")
		user_level=rs("user_level")
		user_logclassid_Nav=rs("logclassid_Nav")'�Ƿ��Ǳ��󶨵���Ŀ�ռ䣻
		'�ж��Ƿ���ʵ����
		if true_domain=1 then'��������˶��������Ĺ��ܣ�
			if rs("custom_domain")<>"" and not isnull(rs("custom_domain")) then
				user_domain=rs("custom_domain")
			end if
			user_truepath="http://"&user_domain&"/"
			log_truepath=""
		else'���û�����ö����������ܣ�
			user_truepath=blogdir&user_path&"/"
			log_truepath=blogdir
		end if
        If user_skin_main = "" Or IsNull(user_skin_main) Then
            Set rstmp = oblog.Execute("select skinmain,skinshowlog from oBlog_userskin where isdefault=1")
            If Not rstmp.EOF Then
                user_skin_main = rstmp(0)
                user_skin_showlog = rstmp(1)
            Else
                Set rstmp = Nothing
                Set rs = Nothing
                Response.Write ("ģ�����")
                Response.End
            End If
        End If
        If user_Blogpassword = "" Or IsNull(user_Blogpassword) Then ispwBlog = False Else ispwBlog = True
    End Property
    
    Public Sub Update_log(logid, resp)
        Dim sql, rstmp
        Dim show_topic, show_emot, show_addtime, show_logtext, show_author, show_loginfo, show_more, show_logcyc, show_topictxt, show_logmore, show, log_month, user_logpath1, log_title, commentasc
        Dim homepage_str, commentid, commenttopic, strtmp, encommment, i, filename, injsfile,user_logpath,logtype
        logid = CLng(logid)'passcheck=0��ʾ��Ҫ����Ա��˺󣬲�����ʾ������
        Set rs = oblog.Execute("select face,topic,logtext,author,istop,isencomment,addtime,ishide,ispassword,isbest,commentnum,trackbacknum,passcheck,authorid,filename,logtype,logPhoto_AddressURL from oblog_log where logid=" & logid)
        If rs.EOF Then Exit Sub
        If Int(Month(rs("addtime"))) < 10 Then
            log_month = Year(rs("addtime")) & "0" & Month(rs("addtime"))
        Else
            log_month = Year(rs("addtime")) & Month(rs("addtime"))
        End If
		if logfilepath=0 then'�����ļ��ķ��õ�ַ��
			user_logpath=user_path
		else
			user_logpath=user_path&"/archives/"&trim(year(rs("addtime")))
		end if
		logtype=rs("logtype")
        filename = Trim(rs("filename"))
        If filename = "" Or IsNull(filename) Then filename = logid
        encommment = rs("isencomment")
        If rs("ishide") = 1 And showpwlog=false Then strtmp = "������Ϊ��������...<a href='" & blogurl & "more.asp?id=" & logid & "'>���������֤ҳ��</a>��"
        If rs("ispassword") <> "" And showpwlog=false Then strtmp = "<form method='post' action='" & blogurl & "more.asp?id=" & logid & "'>���������·������룺<input type=""password"" size=""15"" name=""password"" />  <input type=""submit"" value=""�ύ""></form>"
        If rs("passcheck") = 0 Then strtmp = "��������Ҫ������˺�ſɼ���"
        If user_nickName <> "" Then user_showName = user_nickName Else user_showName = user_userName
        If rs("face") = "0" Then show_emot = "" Else show_emot = "<img src=""" & blogurl & "images/face/" & rs("face") & ".gif"" />"
        show_topictxt = oblog.filt_html(rs("topic"))
		Dim show_log_indexImage'WL;
		If rs("logPhoto_AddressURL")<>"" And len(rs("logPhoto_AddressURL"))>6 Then'WL;
			'''show_log_indexImage = "<img src='"& rs("logPhoto_AddressURL") &"' />"
			show_log_indexImage = ""
		Else'WL;
			show_log_indexImage = ""
		End If'WL;
		
        log_title = show_topictxt
        commenttopic = "Re:" & show_topictxt
        If rs("isbest") = 1 Then show_topictxt = show_topictxt & "��<img src=""" & blogurl & "images/title03.gif"" />"
        show_topic = show_emot
        show_addtime = Year(rs("addtime"))&"-"&Month(rs("addtime"))&"-"&Day(rs("addtime")) &"&nbsp;&nbsp;&nbsp;"&oblog.WeekDateE(rs("addtime"))&""
        show_topic = show_topic & show_topictxt
        If user_nickName = "" Or IsNull(user_nickName) Then
            show_author = user_userName
        Else
            show_author = user_nickName
        End If
        If rs("authorid") <> user_id Then show_author = rs("author")
        'show_loginfo = show_author & " ������ " & show_addtime'WL
		'show_loginfo = "MyHomestay" & "&nbsp;&nbsp; ������ " & show_addtime
		show_loginfo = "&nbsp;&nbsp; ������ " & show_addtime
        show_more = "<a href=""#"" >�Ķ�ȫ��<span id=""ob_logreaded""></span></a>"
        show_more = show_more & " | " & "<a href=""#cmt"">�ظ�(" & rs("commentnum") & ")</a>"
        show_more = show_more & " | <a href=""" & blogurl & "showtb.asp?id=" & logid & """ target=""_blank"">����ͨ��<span id=""ob_tbnum""></span></a>"
        injsfile = "<Script src=""" & blogurl & "count.asp?action=logtb31&id=" & logid & """></Script>"
        show_more = show_more & " | <a href=""" & blogurl & "user_post.asp?logid=" & logid & """ target=""_blank"">�༭</a>"
        If strtmp <> "" Then
            show_logtext = strtmp
        Else
            'show_logtext = oblog.RemoveHTMLTag(rs("logtext"))'WL
			show_logtext = rs("logtext")'WL
            If Left(show_logtext, 7) = "#isubb#" Then
                show_logtext = UBBCode(show_logtext, 1)
                show_logtext = Replace(show_logtext, Chr(10), "<br /> ")
                'show_logtext=oblog.filt_html_b(show_logtext)
            End If
            show_logtext = Replace(show_logtext, "#isubb#", "")
            '''show_logtext = filtimg(show_logtext)
			show_logtext = show_logtext'''WL;��һ�ԣ�
        End If
        show_logcyc = Replace(user_skin_showlog, "$show_topic$", show_topic)
        show_logcyc = Replace(show_logcyc, "$show_loginfo$", show_loginfo)
        show_logcyc = Replace(show_logcyc, "$show_logtext$", show_logtext)
        show_logcyc = Replace(show_logcyc, "$show_more$", show_more)
        show_logcyc = Replace(show_logcyc, "$show_emot$", show_emot)
        show_logcyc = Replace(show_logcyc, "$show_author$", show_author)
        show_logcyc = Replace(show_logcyc, "$show_addtime$", show_addtime)
        show_logcyc = Replace(show_logcyc, "$show_topictxt$", show_topictxt)
		
		show_logcyc = Replace(show_logcyc, "$show_log_indexImage$", show_log_indexImage)'WL;
        If EN_TAGS = 1 Then
            Dim sBlogTag
            sBlogTag = Tags_ShowForBlog(logid,user_truepath)
            'If sBlogTag <> "" Then sBlogTag = P_TAGS_DESC & ":" & sBlogTag
            show_logcyc = Replace(show_logcyc, "$show_blogtag$", sBlogTag)
        Else
            show_logcyc = Replace(show_logcyc, "$show_blogtag$", "")
        End If
        show_logcyc = Replace(show_logcyc, "$show_blogzhai$", "<div id=""blogzhai""></div>")
        show_logmore = show_logcyc
		show_logmore = show_logmore&"<div id=""morelog""><ul>" '��ͷ��
			'(wl)set rstmp=oblog.execute("select top 1 logfile,topic from oblog_log where logid<"&logid&" and userid="&user_id&" and logtype="&logtype&" and isdraft=0 order by addtime desc")
			set rstmp=oblog.execute("select top 1 logfile,topic from oblog_log where logid<"&logid&" and logtype="&logtype&" and isdraft=0 order by addtime desc")
			if not rstmp.eof then
				show_logmore = show_logmore&"<li>��һƪ��"&"<a href="""&log_truepath&rstmp(0)&""">"&oblog.filt_html(rstmp(1))&"</a></li>"
				rstmp.movenext
			end if
			'(wl)set rstmp=oblog.execute("select top 1 logfile,topic from oblog_log where logid>"&logid&" and userid="&user_id&" and logtype="&logtype&" and isdraft=0 order by addtime asc")
			set rstmp=oblog.execute("select top 1 logfile,topic from oblog_log where logid>"&logid&" and logtype="&logtype&" and isdraft=0 order by addtime asc")
			if not rstmp.eof then
				show_logmore = show_logmore&"<li>��һƪ��"&"<a href="""&log_truepath&rstmp(0)&""">"&oblog.filt_html(rstmp(1))&"</a></li>"
				rstmp.movenext
			end if
		show_logmore = show_logmore&"</ul></div>" '��β��
        If strtmp = "" Then
            If user_commentasc = 1 Then commentasc = " order by commentid asc" Else commentasc = " order by commentid desc"
            Set rs = oblog.Execute("select top 40 comment_user,commenttopic,comment,addtime,commentid,homepage,isguest from oblog_comment where mainid=" & logid & commentasc)
            If Not rs.EOF Then
                While Not rs.EOF'ѭ���滻��ǩ����ô���Ϳ���Ų�г�������ۣ�
                    If IsNull(rs(5)) Then
                        homepage_str = "������ҳ"
                    Else
                        If Trim(Replace(rs(5), "http://", "")) = "" Then
                            homepage_str = "������ҳ"
                        Else
                            homepage_str = "<a href=""" & oblog.filt_html(rs(5)) & """ target=""_blank"">������ҳ</a>"
                        End If
                    End If
                    commentid = rs(4)
                    show_topic = oblog.filt_html(rs(1)) & "<a name='" & rs(4) & "'></a>"
                    If rs(6) = 1 Then
                        show_author = "<span id=""n_" & commentid & """>" & oblog.filt_html(rs(0)) & "(�ο�)</span>"
                    Else
                        show_author = "<span id=""n_" & commentid & """>" & oblog.filt_html(rs(0)) & "</span>"
                    End If
                    show_addtime = "<span id=""t_" & commentid & """>" & rs(3) & "</span>"
                    show_topictxt = show_topic
                    show_loginfo = show_author & "����������" & show_addtime
                    show_logtext = "<span id=""c_" & commentid & """>" & oblog.FilterUbbFlash(filtscript(rs(2))) & "</span>"
                    show_more = homepage_str & " | <a href=""javascript:reply_quote('" & commentid & "')"" >����</a> | <a href=""#top"">����</a>"
                    show_more = show_more & " | <a href=""" & blogurl & "user_comments.asp?action=del&id=" & commentid & """  target=""_blank"">ɾ��</a>"
					show_more = show_more & " | <a href=""" & blogurl & "user_comments.asp?action=modify&re=true&id=" & commentid & """  target=""_blank"">�ظ�</a>"
                    show_logcyc = Replace(user_skin_showlog, "$show_topic$", show_topic)
                    show_logcyc = Replace(show_logcyc, "$show_loginfo$", show_loginfo)
                    show_logcyc = Replace(show_logcyc, "$show_logtext$", show_logtext)
                    show_logcyc = Replace(show_logcyc, "$show_more$", show_more)
                    show_logcyc = Replace(show_logcyc, "$show_emot$", "")
                    show_logcyc = Replace(show_logcyc, "$show_author$", show_author)
                    show_logcyc = Replace(show_logcyc, "$show_addtime$", show_addtime)
                    show_logcyc = Replace(show_logcyc, "$show_topictxt$", show_topictxt)
                    show_logmore = show_logmore & show_logcyc
                    show_logmore = Replace(show_logmore, "$show_blogtag$", "")
                    show_logmore = Replace(show_logmore, "$show_blogzhai$", "")
                    rs.movenext
                    i = i + 1
                Wend'ѭ���滻��ǩ����ô���Ϳ���Ų�г�������ۣ�
            End If
			
            If i >= 40 Then
            show_logmore = show_logmore & "<div id=""saveurl""> ::<a href=""" & blogurl & "more.asp?action=comment&id=" & logid & "&page=5"" target=""_blank"">�鿴��������</a>::</div>"
            End If
            If encommment = 1 Then
                Dim strguest
                If oblog.setup(11, 0) = 1 Then strguest = "(�ο�������������)" Else strguest = ""
                show_logmore = filt_inc(show_logmore)
                show_logmore = show_logmore & "#ad_usercomment#<a name='cmt'></a><h2>�������ۣ�</h2>" & vbCrLf
                show_logmore = show_logmore & "<div id=""form_comment""><form action='" & blogurl & "savecomment.asp?logid=" & logid & "' method='post' name='commentform' id='commentform' onSubmit='return Verifycomment()'>" & vbCrLf
                show_logmore = show_logmore & "<ul>�ǳƣ�<input name='UserName' type='text' id='UserName' size='15' maxlength='20' value='' /></ul>" & vbCrLf
                show_logmore = show_logmore & "<ul>���룺<input name='Password' type='password' id='Password' size='15' maxlength='20' value='' /> " & strguest & "</ul>" & vbCrLf
                show_logmore = show_logmore & "<ul>��ҳ��<input name='homepage' type='text' id='homepage' size='42' maxlength='50' value='http://' /></ul>"
                show_logmore = show_logmore & "<ul>���⣺<input name='commenttopic' type='text' id='commenttopic' size='42' maxlength='50' value='" & commenttopic & "' /></ul>"
                show_logmore = show_logmore & "<ul><input type='hidden' name='edit' id='edit' value='' />" & vbCrLf
                show_logmore = show_logmore & "<div id=""oblog_edit"">"&oblog.setup(82,0)&"</div>" & vbCrLf
                show_logmore = show_logmore & "</ul>" & vbCrLf
                show_logmore = show_logmore & "<ul><span id=""ob_code""></span><input type='submit' value=' �ύ '></ul>" & vbCrLf
                show_logmore = show_logmore & "</form></div>" & vbCrLf
            End If
        End If
        show = Replace(user_skin_main, "$show_log$", show_logmore)
        If showpwblog = False And showpwlog = False Then
			 show = repl_label(show, injsfile, "" & BlogName &" --" &log_title, user_userName & "," & user_nickName, log_title, Left(RemoveHTML(show_logtext), 80), log_month)'����log��WL
			if true_domain=1 then
				user_logpath1 = replace(user_logpath,user_path,"http://"&user_domain) & "/" & filename & "." & f_ext
			else
				user_logpath1 = user_logpath & "/" & filename & "." & f_ext
			end if
            show=replace(show,"$show_calendar$","<!-- #include file=""..\..\calendar\"&log_month&".htm"" -->")
            If ispwblog = False Then
                savefile user_logpath,"\"&filename&"."&f_ext,show
            Else
                savefile user_logpath,"\"&filename&"."&f_ext,"<script language=javascript>window.location.replace('"&blogurl&"pwblog.asp?action=log&userid="&user_id&"&logid="&logid&"')</script>"
            End If
            oblog.execute("update oblog_log set logfile='"&user_logpath1&"' where logid="&logid)
            If resp = 1 Then
                gourl = user_logpath1
                response.Write("<li><a href="&user_logpath1&" target=_blank>����鿴���ɵ�"&tname&"</a></li>")
                response.Write("<li><a href=user_post.asp?logid="&logid&"&t="&request("t")&">����޸ĸո��ύ��"&tname&"</a></li>")
            ElseIf resp = 2 Then
                response.Redirect (user_logpath1)
            ElseIf resp = 3 Then
                gourl = user_logpath1
            End If
        Else
            If f_ext = "htm" Or f_ext = "html" Then
                m_log=replace(show,"$show_calendar$","<div id=""calendar""></div><script src='"&user_path&"/calendar/"&log_month&".htm'></script>")
            Else
                m_log=replace(show,"$show_calendar$","<div id=""calendar"">"&oblog.readfile(user_path&"\calendar",log_month&".htm")&"</div>")
            End If
			m_log=m_log&injsfile
        End If
    End Sub
    
    Public Sub showcmt(logid)
        Dim sql, rstmp
        Dim show_topic, show_emot, show_addtime, show_logtext, show_author, show_loginfo, show_more, show_logcyc, show_topictxt, show_logmore, show, commentasc
        Dim homepage_str, commentid, strtmp
        logid = CLng(logid)
        If user_commentasc = 1 Then commentasc = " order by commentid asc" Else commentasc = " order by commentid desc"
        Set rs = Server.CreateObject("Adodb.RecordSet")
        rs.open "select comment_user,commenttopic,comment,addtime,commentid,homepage,isguest from oBlog_comment where mainid=" & logid & commentasc, conn, 1, 1
        If rs.EOF And rs.bof Then
            show_logmore = show_logmore & "����0ƪ����<br>"
          Else
            Dim totalPut, show_page, strFileName, i
            strFileName = "more.asp?action=comment&id=" & logid
            totalPut = rs.recordcount
            If currentPage < 1 Then
                currentPage = 1
            End If
            If (currentPage - 1) * MaxPerPage > totalPut Then
                If (totalPut Mod MaxPerPage) = 0 Then
                    currentPage = totalPut \ MaxPerPage
                  Else
                    currentPage = totalPut \ MaxPerPage + 1
                End If
            End If
            If (currentPage - 1) * MaxPerPage < totalPut Then
                rs.Move (currentPage - 1) * MaxPerPage
                show_page = oblog.showpage(strFileName, totalPut, MaxPerPage, False, True, "ƪ����")
            End If
            Do While Not rs.EOF
                If IsNull(rs(5)) Then
                    homepage_str = "������ҳ"
                  Else
                    If Trim(Replace(rs(5), "http://", "")) = "" Then
                        homepage_str = "������ҳ"
                      Else
                        homepage_str = "<a href=""" & oblog.filt_html(rs(5)) & """ target=""_blank"">������ҳ</a>"
                    End If
                End If
                commentid = rs(4)
                show_topic = oblog.filt_html(rs(1)) & "<a Name='" & rs(4) & "'></a>"
                If rs(6) = 1 Then
                    show_author = "<span id=""n_" & commentid & """>" & oblog.filt_html(rs(0)) & "(�ο�)</span>"
                  Else
                    show_author = "<span id=""n_" & commentid & """>" & oblog.filt_html(rs(0)) & "</span>"
                End If
                show_addtime = "<span id=""t_" & commentid & """>" & rs(3) & "</span>"
                show_topictxt = show_topic
                show_loginfo = show_author & "����������" & show_addtime
                show_logtext = "<span id=""c_" & commentid & """>" &  oblog.FilterUbbFlash(filtscript(rs(2))) & "</span>"
                show_more = homepage_str & " | <a href=""javascript:reply_quote('" & commentid & "')"" >����</a> | <a href=""#top"">����</a>"
                show_more = show_more & " | <a href=""user_comments.asp?action=del&id=" & commentid & """  target=""_blank"">ɾ��</a>"
                show_logcyc = Replace(user_skin_showlog, "$show_topic$", show_topic)
                show_logcyc = Replace(show_logcyc, "$show_loginfo$", show_loginfo)
                show_logcyc = Replace(show_logcyc, "$show_logtext$", show_logtext)
                show_logcyc = Replace(show_logcyc, "$show_more$", show_more)
                show_logcyc = Replace(show_logcyc, "$show_emot$", "")
                show_logcyc = Replace(show_logcyc, "$show_author$", show_author)
                show_logcyc = Replace(show_logcyc, "$show_addtime$", show_addtime)
                show_logcyc = Replace(show_logcyc, "$show_topictxt$", show_topictxt)
                show_logmore = show_logmore & show_logcyc
                i = i + 1
                If i >= MaxPerPage Then Exit Do
                rs.movenext
            Loop
        End If
		show_logmore = Replace(show_logmore, "$show_blogtag$", "")
        show_logmore = Replace(show_logmore, "$show_blogzhai$", "")
        show_logmore = show_logmore & show_page
        If encommment = 1 Then
            Dim strguest
            If oblog.setup(11, 0) = 1 Then strguest = "(�ο�������������)" Else strguest = ""
            show_logmore = filt_inc(show_logmore)
            show_logmore = show_logmore & "#ad_usercomment#<a Name='comment'></a><h2>�������ۣ�</h2>" & vbCrLf
            show_logmore = show_logmore & "<div id=""form_comment""><form action='" & blogurl & "savecomment.asp?logid=" & logid & "' method='post' name='commentform' id='commentform' onSubmit='return Verifycomment()'>" & vbCrLf
            show_logmore = show_logmore & "<ul>�ǳƣ�<input Name='UserName' type='text' id='UserName' size='15' maxlength='20' value='' /></ul>" & vbCrLf
            show_logmore = show_logmore & "<ul>���룺<input Name='Password' type='password' id='Password' size='15' maxlength='20' value='' /> " & strguest & "</ul>" & vbCrLf
            show_logmore = show_logmore & "<ul>��ҳ��<input Name='homepage' type='text' id='homepage' size='42' maxlength='50' value='http://' /></ul>"
            show_logmore = show_logmore & "<ul>���⣺<input Name='commenttopic' type='text' id='commenttopic' size='42' maxlength='50' value='" & commenttopic & "' /></ul>"
            show_logmore = show_logmore & "<ul><input type='hidden' Name='edit' id='edit' value='' />" & vbCrLf
            show_logmore = show_logmore & "<div id=""oblog_edit""></div>" & vbCrLf
            show_logmore = show_logmore & "</ul>" & vbCrLf
            show_logmore = show_logmore & "<ul><script src=""" & blogurl & "count.asp?action=code""></script><input type='Submit' value=' �ύ '></ul>" & vbCrLf
            show_logmore = show_logmore & "</form></div>" & vbCrLf
        End If
        show = Replace(user_skin_main, "$show_log$", show_logmore)
        If f_ext = "htm" Or f_ext = "html" Then
            m_commentsmore = Replace(show, "$show_calendar$", "<div id=""calendar""></div><script src='" & user_path & "\calendar\" & log_month & ".htm'></script>")
        ElseIf Page="cmd" Then
        	m_commentsmore = Replace(show, "$show_calendar$", "<div id=""calendar""></div><" & "%'=Calendar(intYear,intMonth,intDay)%" & ">")        	
    	Else
            m_commentsmore = Replace(show, "$show_calendar$", "<div id=""calendar"">" & oblog.readfile(user_path & "\calendar", log_month & ".htm") & "</div>")
        End If
    End Sub
    
    Public Sub Update_index(resp)
        Dim show_topic, show_emot, show_addtime, show_logtext, show_author, show_loginfo, show_more, show_logcyc, show_topictxt, show_logmore, show, strtmp, xmlstr, rstmp, strart, i, strlogn,  injsfile,substr
		
        xmlstr = "<?xml version=""1.0"" encoding=""GB2312""?>" & vbCrLf & "<rss version=""2.0"">" & vbCrLf
        xmlstr = xmlstr & "<channel>" & vbCrLf & "<title><![CDATA[" & BlogName & "]]></title>" & vbCrLf
		if true_domain=1 then
			xmlstr = xmlstr & "<link>http://" & user_domain & "/index." & f_ext & "</link>" & vbCrLf
		else
			xmlstr = xmlstr & "<link>" & Trim(oblog.setup(3, 0)) & user_path & "/index." & f_ext & "</link>" & vbCrLf
		end if
        xmlstr = xmlstr & "<description><![CDATA[" & BlogName & "]]></description>" & vbCrLf
        If user_showlog_num = 0 Then user_showlog_num = 1
		Set rs = oblog.execute("select subjectid,subjectname from oblog_subject where userid="&user_id)
        While Not rs.EOF
            substr = substr & rs(0) & "!!??((" & rs(1) & "##))=="
            rs.movenext
        Wend
        Set rs = oblog.Execute("select top " & user_showlog_num & " face,topic,subjectid,logid,istop,addtime,ishide,commentnum,showword,ispassword,iis,trackbacknum,isbest,blog_password,author,logfile,ishide,authorid,passcheck,Abstract,logtext,logPhoto_AddressURL from oblog_log where userid=" & user_id & " and passcheck=1 and isdraft=0 and logtype=0 order by istop desc,addtime desc")
		
		Dim show_log_indexImage'WL;
		
			
        While Not rs.EOF
		
		If rs("logPhoto_AddressURL")<>"" And len(rs("logPhoto_AddressURL"))>6 Then'WL;
			show_log_indexImage = "<a target=_blank href='" & rs("logPhoto_AddressURL") &"'>"
			show_log_indexImage = show_log_indexImage & "<img src='"& rs("logPhoto_AddressURL") &"' />"
			show_log_indexImage = show_log_indexImage & "</a>"
		Else'WL;
			show_log_indexImage = ""
		End If'WL;
			strtmp=""
            If rs("face") = "0" Then show_emot = "" Else show_emot = "<img src=" & blogurl & "images/face/" & rs("face") & ".gif >"
            If user_nickName = "" Or IsNull(user_nickName) Then
                show_author = user_userName
            Else
                show_author = user_nickName
            End If
            If rs("authorid") <> user_id Then show_author = rs("author")
            show_addtime = Year(rs("addtime"))&"-"&Month(rs("addtime"))&"-"&Day(rs("addtime")) &"&nbsp;&nbsp;&nbsp;"&oblog.WeekDateE(rs("addtime"))&""
            show_topic = show_emot
            If rs("istop") = 1 Then show_topic = "[�ö�]"
            If rs("subjectid") > 0 Then
                    show_topic = show_topic & "<a href=""" & user_truepath&"MyHomestay."&f_ext&"?uid="&user_id&"&do=blogs&id=" & rs("subjectid") & """>[" & oblog.filt_html(getsubname(substr,rs("subjectid"))) & "]</a>"
            End If
            show_topictxt = "<a target=_blank href=""" & log_truepath & rs("logfile") & """>" & oblog.filt_html(rs("topic")) & "</a>"
            If rs(12) = 1 Then show_topictxt = show_topictxt & "��<img src=" & blogurl & "images/title03.gif >"
            show_topic = show_topic & show_topictxt
            If rs("istop") = 1 Then show_topictxt = "[�ö�]" & show_topictxt
            'show_loginfo = show_author & " ������ " & show_addtime
			show_loginfo =  " ������ " & show_addtime
            show_more = "<a href=""" & log_truepath & rs("logfile") & """>�Ķ�ȫ��<span id=""ob_logr" & rs("logid") & """></span></a>"
            show_more = show_more & " | <a href=""" & log_truepath & rs("logfile") & "#cmt"">�ظ�<span id=""ob_logc" & rs("logid") & """></span></a>"
            show_more = show_more & " | <a href=""" & blogurl & "showtb.asp?id=" & rs(3) & """ target=""_blank"">����ͨ��<span id=""ob_logt" & rs("logid") & """></span></a>"
            '----------------------------------------
            'ȡ������ժҪ����
            If IsNull(rs("Abstract")) Or Trim(rs("Abstract")) = "" or rs("ishide")=1 or rs("ispassword") <> "" or rs("passcheck") = 0  Then
                '������ǰ����
                If rs("ishide") = 1 Then strtmp = "������Ϊ�������£������ѿɼ���<a href='" & blogurl & "more.asp?id=" & rs("logid") & "'>���������֤ҳ��</a>��"
                If rs("ispassword") <> "" Then strtmp = "<form method='post' action='" & blogurl & "more.asp?id=" & rs("logid") & "' target='_blank'>���������·������룺<input type=""password"" size=""15"" name=""password"" />  <input type=""submit"" value=""�ύ""></form>"
                If rs("passcheck") = 0 Then strtmp = "��������Ҫ����Ա��˺�ſɼ���"
                If strtmp <> "" Then
                    show_logtext = strtmp
                Else
                    'show_logtext=filtscript(rs(6))
                    show_logtext = oblog.nohtml(rs("logtext"))'WL
                    show_logtext = trimlog(show_logtext, rs("showword"))
                    If Left(show_logtext, 7) = "#isubb#" Then
                        show_logtext = UBBCode(show_logtext, 1)
                        show_logtext = Replace(show_logtext, Chr(10), "<br /> ")
                    'show_logtext=oblog.filt_html_b(show_logtext)
                    End If
                    show_logtext = Replace(show_logtext, "#isubb#", "")
                    show_logtext = filtimg(show_logtext)
                    If oblog.setup(29, 0) = 1 Then show_logtext = profilthtm(show_logtext)
                End If
            Else
                show_logtext = Left(oblog.nohtml(rs("logtext")),230)'WL
            End If
			show_logtext=oblog.filt_badword(show_logtext)
            '----------------------------------------
            strlogn = strlogn & "$" & rs("logid")
            'ֻ�Ǽǿ��ż�¼
            If (rs("ispassword")="" Or IsNull(rs("ispassword"))) And rs("blog_password")=0 And rs("ishide")=0 Then
	            xmlstr = xmlstr & "<item>" & vbCrLf & "<title><![CDATA[" & rs("topic") & "]]></title>" & vbCrLf
				if true_domain=1 then
					xmlstr = xmlstr & "<link>" &rs("logfile") & "</link>" & vbCrLf
				else
	            	xmlstr = xmlstr & "<link>" & Trim(oblog.setup(3, 0)) & rs("logfile") & "</link>" & vbCrLf
				end if
	            xmlstr = xmlstr & "<description><![CDATA[" & oblog.trueurl(show_logtext) & "]]></description>" & vbCrLf
	            xmlstr = xmlstr & "<author>" & show_author & "</author>" & vbCrLf
	            xmlstr = xmlstr & "<pubDate>" & show_addtime & "</pubDate>" & vbCrLf & "</item>" & vbCrLf
	        End If
            show_logcyc = Replace(user_skin_showlog, "$show_topic$", show_topic)
            show_logcyc = Replace(show_logcyc, "$show_loginfo$", show_loginfo)
            show_logcyc = Replace(show_logcyc, "$show_logtext$", show_logtext)
			
			show_logcyc = Replace(show_logcyc, "$show_log_indexImage$", show_log_indexImage)
			
            show_logcyc = Replace(show_logcyc, "$show_more$", show_more)
            show_logcyc = Replace(show_logcyc, "$show_emot$", show_emot)
            show_logcyc = Replace(show_logcyc, "$show_author$", show_author)
            show_logcyc = Replace(show_logcyc, "$show_addtime$", show_addtime)
            show_logcyc = Replace(show_logcyc, "$show_topictxt$", show_topictxt)
            rs.movenext
            show_logmore = show_logmore & show_logcyc
        Wend
        xmlstr = xmlstr & vbCrLf & "</channel>" & vbCrLf & "</rss>"
		set rstmp=oblog.execute("select count(logid) from oblog_log where userid=" & user_id & " and passcheck=1 and isdraft=0")
		CurrentPage=1
        strart =  CreateStaticPageBar(rstmp(0),user_showlog_num,0) 
        injsfile = "<script src=""" & blogurl & "count.asp?action=logs&id=" & strlogn & """></script>"
        'show = Replace(user_skin_main, "$show_log$", show_logmore & strart &"</div>")
		'show = Replace(user_skin_main, "$show_log$", show_logmore & strart)WL
		show = Replace(user_skin_main_Nav, "$show_log$", show_logmore & strart)
        show = Replace(show, "$show_blogtag$", "")
        show = Replace(show, "$show_blogzhai$", "")
        show = filt_inc(show)
        If showpwblog = False Then
            show = repl_label(show, injsfile, BlogName, user_userName & "," & user_nickName, BlogName, user_siteinfo, newcalendar(blogdir&user_path&"/" & "calendar"))
            If ispwblog = False Then
                savefile user_path, "\index." & f_ext, Show
                savefile user_path, "\rss2.xml", xmlstr
            Else
                savefile user_path, "\rss2.xml", "��blog�Ѽ���"
                savefile user_path,"\index."&f_ext,"<script language=javascript>window.location.replace('"&blogurl&"pwblog.asp?action=blog&userid="&user_id&"')</script>"
            End If
            If resp = 1 Then
                response.Write("<li><a href="&user_path&"/index."&f_ext&" target=_blank>����鿴���ɵ���ҳ</a></li>")
            ElseIf resp = 2 Then
                response.Redirect(user_path&"/index."&f_ext)
            End If
        Else
            m_index = show&injsfile
        End If
    End Sub
    
        
    Public Sub Update_message(resp)
        Dim show_topic, show_emot, show_addtime, show_logtext, show_author, show_loginfo, show_more, show_logcyc, show_topictxt, show_logmore, show
        Dim homepage_str, user_filepath,strPageBar,lngAll
        Set rs = oblog.Execute("select count(*) from oblog_message where userid=" & user_id)
        lngAll=rs(0)
        Set rs = oblog.Execute("select top " & user_showlog_num & " message_user,messagetopic,message,addtime,messageid,homepage,isguest from oblog_message where userid=" & user_id & " order by messageid desc")
        If Not rs.EOF Then
            While Not rs.EOF
                If IsNull(rs(5)) Then
                    homepage_str = "������ҳ"
                Else
                    If Trim(Replace(rs(5), "http://", "")) = "" Then
                        homepage_str = "������ҳ"
                    Else
                        homepage_str = "<a href=""" & oblog.filt_html(rs(5)) & """ target=""_blank"">������ҳ</a>"
                    End If
                End If
                show_topic = oblog.filt_html(rs(1)) & "<a name='" & rs(4) & "'></a>"
                If rs(6) = 1 Then
                    show_author = oblog.filt_html(rs(0)) & "(�ο�)"
                Else
                    show_author = oblog.filt_html(rs(0))
                End If
                show_addtime = rs(3)
                show_topictxt = show_topic
                show_loginfo = show_author & "����������" & show_addtime
                show_logtext = oblog.FilterUbbFlash(filtscript(rs(2)))
                show_more = homepage_str & " | <a href='#cmt'>ǩд����</a> | <a href='"&blogurl&"user_messages.asp?action=modify&re=true&id=" & rs(4) & "'>�ظ�</a>"
                show_more = show_more & " | <a href=""" & blogurl & "user_messages.asp?action=del&id=" & rs(4) & """  target=""_blank"">ɾ��</a>"
                show_logcyc = Replace(user_skin_showlog, "$show_topic$", show_topic)'ÿһ����¼�����ԣ����ʹ���һ��showlogģ�壬������ʾ��wl��
                show_logcyc = Replace(show_logcyc, "$show_loginfo$", show_loginfo)
                show_logcyc = Replace(show_logcyc, "$show_logtext$", show_logtext)
                show_logcyc = Replace(show_logcyc, "$show_more$", show_more)
                show_logcyc = Replace(show_logcyc, "$show_emot$", "")
                show_logcyc = Replace(show_logcyc, "$show_author$", show_author)
                show_logcyc = Replace(show_logcyc, "$show_addtime$", show_addtime)
                show_logcyc = Replace(show_logcyc, "$show_topictxt$", show_topictxt)
                show_logcyc = Replace(show_logcyc, "$show_blogtag$", "")
                show_logcyc = Replace(show_logcyc, "$show_blogzhai$", "")
                show_logmore = show_logmore & show_logcyc'��show_logmoreѭ����¼�����еļ�¼�����ԣ�wl��
                rs.movenext
            Wend
            strPageBar = CreateStaticPageBar(lngAll,user_showlog_num,1) '������ҳ������
        Else
            show_logmore = "��������"'��show_logmore��ʾû�м�¼�����ԣ�wl��
            strPageBar =""'��ҳ����Ϊ�գ�
        End If
        show_logmore = show_logmore &  strPageBar
        Dim strguest, strart, i
        If oblog.setup(11, 0) = 1 Then strguest = "(�ο�������������)" Else strguest = ""
        show_logmore = filt_inc(show_logmore)
        show_logmore = show_logmore & strart & "#ad_usercomment#<a name='cmt'></a><h2>ǩд���ԣ�</h2>" & vbCrLf
        show_logmore = show_logmore & "<div id=""form_comment""><form action='" & blogurl & "savemessage.asp?userid=" & user_id & "' method='post' name='commentform' id='commentform' onSubmit='return Verifycomment()'>" & vbCrLf
        show_logmore = show_logmore & "<ul>�ǳƣ�<input name='UserName' type='text' id='UserName' size='15' maxlength='20' value='' /></ul>" & vbCrLf
        show_logmore = show_logmore & "<ul>���룺<input name='Password' type='password' id='Password' size='15' maxlength='20' value='' /> " & strguest & "</ul>" & vbCrLf
        show_logmore = show_logmore & "<ul>��ҳ��<input name='homepage' type='text' id='homepage' size='42' maxlength='50' value='http://' /></ul>"
        show_logmore = show_logmore & "<ul>���⣺<input name='commenttopic' type='text' id='commenttopic' size='42' maxlength='50' value='' /></ul>"
        show_logmore = show_logmore & "<ul><input type='hidden' name='edit' id='edit' value='' />" & vbCrLf
        show_logmore = show_logmore & "<div id=""oblog_edit""></div>" & vbCrLf
        show_logmore = show_logmore & "</ul>" & vbCrLf
        show_logmore = show_logmore & "<ul><span id=""ob_code""></span><input name='login' type='submit' id='Login' value=' �ύ ' /> " & vbCrLf
        show_logmore = show_logmore & "</ul>" & vbCrLf
        show_logmore = show_logmore & "</form></div>" & vbCrLf
        show_logmore = "<h1>���԰���ҳ(<a href='#cmt'>ǩд����</a>)</h1>" & vbCrLf & show_logmore

        show = Replace(user_skin_main, "$show_log$", show_logmore)'wl��Ҫ������show_logmore�滻��ģ���$show_log$��ǩ����
        if showpwblog=false then
	        show = repl_label(show, "", BlogName & "--���԰�", user_userName & "," & user_nickName, BlogName, BlogName, newcalendar(blogdir&user_path & "/calendar"))'show�������Ѻ���ģ������show����repl_label�滻�������ñ�ǩ���������ض��Ĺ��ܣ���
			if true_domain=1 then
				user_filepath = "http://"&user_domain & "/message." & f_ext
			else
				user_filepath = user_path & "/message." & f_ext
			end if
	        If ispwBlog = False Then'������ܸ�û�����룻
	            savefile user_path, "\message." & f_ext, show
	        Else'���������,����һ���پ�̬ҳmessage.html����ת��̬��֤ҳ�棻
	            savefile user_path, "\message." & f_ext, "<script language=javascript>window.location.replace('" & blogurl & "pwblog.asp?action=message&userid=" & user_id & "')</script>"
	        End If
	        If resp = 1 Then
	            Response.Write ("<li><a href=" & user_filepath & " target=_blank>����鿴���԰�!</a></li>")
	        ElseIf resp = 2 Then
	            Response.Redirect (user_filepath)
	        ElseIf resp = 3 Then
	            GoUrl = user_filepath
	        End If
	    Else
	    	m_message=show
	    End If
    End Sub
            
     Public Sub Update_info(userid)
        Dim show
        show = "<ul>" & vbCrLf
        show = show & "<li>��������:" & user_log_count & "</li>" & vbCrLf
        show = show & "<li>��������:" & user_comment_count & "</li>" & vbCrLf
        show = show & "<li>��������:" & user_message_count & "</li>" & vbCrLf
        show = show & "<li>���ʴ���:<span id=""site_count""></span></li>" & vbCrLf
        show = show & "<li><a href=""" & blogurl & "user_friends.asp?action=saveadd&friendname=" & user_userName & """ target=""_blank"">��Ϊ����</a>��<a href=""javascript:openScript('" & blogurl & "user_pm.asp?action=send&incept=" & user_userName & "',450,400)"">���Ͷ���</a></li>"
        show = show & "</ul>" & vbCrLf
        if showpwblog or showpwlog then m_info=show : exit sub
        savefile user_path, "\inc\show_info.htm", show
    End Sub
    
    Public Sub Update_placard(userid)
        Dim show
        show = user_placard
        show = filtSkinPath(filt_inc(show))
        if showpwblog or showpwlog then m_placard=show : exit sub
        savefile user_path, "\inc\show_placard.htm", show
    End Sub
    
    Public Sub Update_links(userid)
        Dim show
        show = user_links & vbCrLf
        show = filtSkinPath(filt_inc(show))
        if showpwblog or showpwlog then m_links=show : exit sub
        savefile user_path, "\inc\show_links.htm", show
    End Sub
    
    Public Sub Update_newblog(userid)
        Dim n, show,fname
        n = CLng(user_shownewlog_num)
        Set rs = oblog.Execute("select top " & n & " topic,addtime,logfile,logid from [oblog_log] where userid=" & userid & " and isdraft=0 and passcheck=1 order by addtime desc")
        If Not rs.EOF Then show = "<ul>" & vbCrLf
        While Not rs.EOF
            show = show & "<li><a href=""" & log_truepath&rs(2) & """ title=""������" & rs(1) & """>" & oblog.filt_html(Left(rs(0), 18)) & "</a></li>" & vbCrLf
            rs.movenext
            If rs.EOF Then show = show & "</ul>" & vbCrLf
        Wend
        if showpwblog or showpwlog then m_newblog=show : exit sub
        savefile user_path, "\inc\show_newblog.htm", show
    End Sub
    
    Public Sub Update_newmessage(userid)'����ͨ�ð����ļ���
        Dim n, show, userdir, ustr
        n = CLng(user_shownewmessage_num)
        show = "<ul><li><a href="""&user_truepath&"message." & f_ext & "#cmt"">::ǩд����::</a></li>"
        Set rs = oblog.Execute("select top " & n & " user_dir,messagetopic,oblog_message.addtime,message_user,messageid,messagefile from oblog_user,oblog_message where oblog_message.userid=" & userid & " and oblog_user.userid=oblog_message.userid order by messageid desc")
        While Not rs.EOF
            ustr = user_truepath&"message." & f_ext & "#" & rs("messageid")
            show = show & "<li><a href=""" & ustr & """ title=""" & oblog.filt_html(rs("message_user")) & "������" & rs("addtime") & """ >" & oblog.filt_html(Left(rs("messagetopic"), 18)) & "</a></li>" & vbCrLf
            rs.movenext
        Wend
        show = show & "</ul>" & vbCrLf
        if showpwblog or showpwlog then m_newmessage=show : exit sub
        savefile user_path, "\inc\show_newmessage.htm", show
    End Sub
    
    Public Sub Update_comment(userid)
        Dim n, show
        n = CLng(user_shownewcomment_num)
        Set rs = oblog.Execute("select top " & n & " oblog_comment.commenttopic,oblog_comment.addtime,oblog_comment.comment_user,oblog_comment.commentid,oblog_log.logfile from oblog_log,oblog_comment where oblog_comment.mainid=oblog_log.logid and oblog_comment.userid=" & userid & " order by commentid desc")
        If Not rs.EOF Then show = "<ul>" & vbCrLf
        While Not rs.EOF
            show = show & "<li><a href=""" &log_truepath&rs("logfile")& "#" & rs("commentid") & """ title=""" & oblog.filt_html(rs("comment_user")) & "������" & rs("addtime") & """>" & oblog.filt_html(Left(rs("commenttopic"), 18)) & "</a></li>" & vbCrLf
            rs.movenext
            If rs.EOF Then show = show & "</ul>" & vbCrLf
        Wend
        if showpwblog or showpwlog then m_comment=show : exit sub
        savefile user_path, "\inc\show_comment.htm", show
    End Sub
    
    '�����û������·���
    Public Sub Update_Subject(userid)
        Dim n, show
        show = "<ul>" & vbCrLf & "<li><a href=""" & user_truepath&"index." & f_ext & """>��ҳ</a>"
        'show = show & "<li><a href=""" & blogdir & "HomestayBackctrl-index.asp"" target=""blank"">����</a></li>"
        If en_photo = 1 Then
            show = show & vbCrLf & " <a href="""&user_truepath&"MyHomestay." & f_ext &"?uid="&user_id&"&do=album"">���</a> "
        End If
        If EN_TAGS = 1 Then
            show = show & vbCrLf & " <a href="""&user_truepath&"MyHomestay."&f_ext&"?uid=" & user_id &"&do=tags"">��ǩ</a>"
        End If
		show=show&"</li>"
        Set rs = oblog.Execute("select Subjectid,SubjectName,Subjectlognum from oBlog_Subject where  userid=" & userid & " and Subjecttype=0 order by ordernum")
        While Not rs.EOF
            show = show & "<li><a href=""" & user_truepath & "MyHomestay."&f_ext&"?do=blogs&id=" & rs("Subjectid") & "&uid="&user_id&""">" & oblog.filt_html(rs("SubjectName")) & "(" & rs("Subjectlognum") & ")" & "</a></li>" & vbCrLf
            rs.movenext
        Wend
        show = show & "</ul>" & vbCrLf
        'show1 = Replace(show, "<div id=""subject"">", "<div id=""subject_l"">")
        if showpwblog or showpwlog then m_subject=show:m_subject_l=show : exit sub
        savefile user_path, "\inc\show_subject.htm", show
        'ʹ��һ���ļ����ò�ͬ��div id���Ƹ�ʽsavefile user_path, "\inc\show_subject_l.htm", show1
    End Sub
    
    Public Sub Update_calendar(logid)
        Dim c_year, c_month, c_day, logdate, today, tomonth, toyear, sql, s, count, b, c
        Dim thismonth, thisdate, thisyear, startspace, NextMonth, NextYear, PreMonth, PreYear, linkTrue
        Dim linkdays, selectdate, linkcount, ccode
        Dim CommondFile
		CommondFile= user_truepath&"MyHomestay."&f_ext&"?uid="&user_id&"&do=month&month="
        ReDim linkdays(2, 0)
        Set rs = oblog.Execute("select addtime from oBlog_log where oBlog_log.logid=" & CLng(logid))
        If rs.EOF Then Exit Sub
        selectdate = rs(0)
        c_year = CInt(Year(selectdate))
        c_month = CInt(Month(selectdate))
        c_day = CInt(Day(selectdate))
        logdate = c_year & "-" & c_month
        If is_sqldata Then
            Dim cmd, rs
            Set cmd = Server.CreateObject("ADODB.Command")
            Set cmd.ActiveConnection = conn
            cmd.CommandText = "ob_calendar"
            cmd.CommandType = 4
            cmd("@logdate") = logdate
            cmd("@userid") = user_id
            Set rs = cmd.Execute
            Set cmd = Nothing
          Else
            sql = "SELECT addtime,logfile from oBlog_log WHERE datediff('n','" & logdate & "',addtime)>0 and userid=" & user_id
            Set rs = oblog.Execute(sql)
        End If
        Dim theday
        theday = 0
    
        Do While Not rs.EOF
            If Day(rs("addtime")) <> theday Then
                theday = Day(rs("addtime"))
                ReDim Preserve linkdays(2, linkcount)
                linkdays(0, linkcount) = Month(rs("addtime"))
                linkdays(1, linkcount) = Day(rs("addtime"))
                'linkdays(2, linkcount) = blogdir & rs("logfile")
                linkdays(2, linkcount)=user_truepath&"MyHomestay."&f_ext&"?uid="&user_id&"&do=day&day=" & CStr(CDate(Year(rs("addtime")) & "-" & Month(rs("addtime")) & "-" & Day(rs("addtime"))))
                linkcount = linkcount + 1
            End If
            rs.movenext
        Loop
        Set rs = Nothing
        Dim mdays(12)
        mdays(0) = ""
        mdays(1) = 31
        mdays(2) = 28
        mdays(3) = 31
        mdays(4) = 30
        mdays(5) = 31
        mdays(6) = 30
        mdays(7) = 31
        mdays(8) = 31
        mdays(9) = 30
        mdays(10) = 31
        mdays(11) = 30
        mdays(12) = 31
        '�����������
        today = Day(ServerDate(Now()))
        tomonth = Month(ServerDate(Now()))
        toyear = Year(ServerDate(Now()))
        'ָ���������ռ�����
        thismonth = c_month
        thisdate = c_day
        thisyear = c_year
        If IsDate("February 29, " & thisyear) Then mdays(2) = 29
        'ȷ������1�ŵ�����
        startspace = Weekday(thismonth & "-1-" & thisyear) - 1
        NextMonth = c_month + 1
        NextYear = c_year+1
        If NextMonth > 12 Then
            NextMonth = 1
            NextYear = NextYear + 1
        End If
        PreMonth = c_month - 1
        PreYear = c_year-1
        If PreMonth < 1 Then
            PreMonth = 12
            'PreYear = PreYear - 1
        End If
        ccode = "<table width='100%'>" & vbCrLf
        'ccode = ccode & "<caption>" & mName(thismonth) & thisyear & "</caption><tr>" & vbCrLf
        ccode = ccode & "<caption><a href="""& CommondFile & (PreYear & Right("0" & c_month,2)) &""" title=""��һ��""><span class=""arrow""><<</span></a>&nbsp;&nbsp;<a href=""" & CommondFile & PreYear& Right("0" & preMonth,2)&""" title=""��һ��""><span class=""arrow""><</span></a>&nbsp;"& toyear &"<a href=""" & CommondFile & Year(ServerDate(Date)) & Right("0" & Month(ServerDate(Date)),2) & """ title=""���ص���""> - </a>"& c_month&"&nbsp;<a href="""& CommondFile & c_year& Right("0" & NextMonth,2) &""" title=""��һ��""><span class=""arrow"">></span></a>&nbsp;&nbsp;<a href=""" & CommondFile & NextYear & Right("0" & c_month,2) &""" title=""��һ��""><span class=""arrow"">>></span></a></caption><tr>"
        'ccode = ccode & "<caption><a href=""" & CommondFile & c_year& Right("0" & preMonth,2)&""" title=""��һ��""><span class=""arrow""><</span></a> "& toyear &"  <a href=""" & CommondFile & Year(ServerDate(Date)) & Right("0" & Month(ServerDate(Date)),2) & """ title=""���ص���"">-</a> "& c_month&" <a href="""& CommondFile & c_year& Right("0" & NextMonth,2) &""" title=""��һ��""><span class=""arrow"">></span></a></caption><tr>"
        ccode = ccode & "<th>��</th>" & vbCrLf
        ccode = ccode & "<th>һ</th>" & vbCrLf
        ccode = ccode & "<th>��</th>" & vbCrLf
        ccode = ccode & "<th>��</th>" & vbCrLf
        ccode = ccode & "<th>��</th>" & vbCrLf
        ccode = ccode & "<th>��</th>" & vbCrLf
        ccode = ccode & "<th>��</th></tr><tr>" & vbCrLf
        For s = 0 To startspace - 1
            ccode = ccode & "<td align=""center"">&nbsp;</td>" & vbCrLf
        Next
        count = 1
        While count <= mdays(thismonth)
            For b = startspace To 6
                ccode = ccode & "<td align=center>"
                linkTrue = "False"
                For c = 0 To UBound(linkdays, 2)
                    If linkdays(0, c) <> "" Then
                        If linkdays(0, c) = thismonth And linkdays(1, c) = count Then
                            ccode = ccode & "<a href='" & linkdays(2, c) & "'>"
                            linkTrue = "True"
                        End If
                    End If
                Next
                If count <= mdays(thismonth) Then ccode = ccode & count
                If linkTrue = "True" Then ccode = ccode & "</a>"
                ccode = ccode & "</td>" & vbCrLf
                count = count + 1
            Next
            If count > mdays(thismonth) Then
                ccode = ccode & "</tr>" & vbCrLf
              Else
                ccode = ccode & "</tr><tr>" & vbCrLf
            End If
            startspace = 0
        Wend
        ccode = ccode & "</table>" & vbCrLf
        If CLng(c_month) < 10 Then c_month = "0" & c_month
        'ccode = "<div id=""calendar"">" & ccode & "</div>"
        savefile user_path, "\calendar\" & c_year & c_month & ".htm", ccode
    End Sub
    
    
    Public Function filt_pwblog(show, log_title)
        update_info (user_id)
        update_subject (user_id)
        update_newblog (user_id)
        update_newmessage (user_id)
        update_links (user_id)
        update_comment (user_id)
		Update_placard (user_id)
        'show=replace(show,"$show_calendar$",oblog.readfile(user_path&"\calendar",log_month&".htm"))
        show = Replace(show, "$show_placard$", "<div id=""placard"">"&m_placard&"</div>")
        show = Replace(show, "$show_subject$", "<div id=""subject"">"&m_subject&"</div>")
        show = Replace(show, "$show_subject_l$", "<div id=""subject_l"">"&m_subject&"</div>")
        show = Replace(show, "$show_newblog$", "<div id=""newblog"">"&m_newblog&"</div>")
        show = Replace(show, "$show_comment$", "<div id=""comment"">"&m_comment&"</div>")
        show = Replace(show, "$show_newmessage$", "<div id=""newmessage"">"&m_newmessage&"</div>")
        show = Replace(show, "$show_links$", "<div id=""links"">"&m_links&"</div><div id=""ad_userlinks""></div>")
        show = Replace(show, "$show_info$", "<div id=""info"">"&m_info&"</div>")
        show = Replace(show, "$show_blogname$", blogname)
		show = Replace(show, "#ad_usercomment#", "<div id=""ad_usercomment""></div>")
		show=replace(show,"$show_xml$","<div id=""xml""><a href="""&user_truepath&"rss2.xml"" target=""_blank""><img src='" & blogurl & "images/xml.gif' width='36' height='14' border='0' /></a></div>")
		
		Call Nav_indexshow_Nav(show)'WL;
		
        If f_ext = "htm" Or f_ext = "html" Then
            show=replace(show,"$show_search$","<div id=""search""></div><script src="""&blogdir&user_path&"/inc/show_search.htm""></script>")
            show=replace(show,"$show_login$","<div id=""ob_login""></div><script src="""&blogdir&user_path&"/inc/show_login.htm""></script>")
			show = Replace(show, "$show_userlogin$", "document.write('<div id=""ob_login""></div><script src="&js_blogurl&"login.asp?action=showjs&injs=1></script>');")
			show = Replace(show, "$show_userlogin_team$", "<div id='ob_login_team'></div><script src='"&js_blogurl&"login.asp?action=showjs_team&injs=1'></script>")
			
        Else
            show=replace(show,"$show_search$",oblog.readfile(user_path&"\inc","show_search.htm"))
            show=replace(show,"$show_login$",oblog.readfile(user_path&"\inc","show_login.htm"))
            'show=replace(show,"$show_xml$",oblog.readfile(user_path&"\inc","show_xml.htm"))
        End If
        show="<script src=""inc/main.js"" type=""text/javascript""></script>"&VbCrLf&show
        'show="<link href=""OblogStyle/OblogUserDefault31.css"" rel=""stylesheet"" type=""text/css"" />"&VbCrLf&"</head>"&VbCrLf&"<body><div id=""ad_usertop""></div>"&show
		show=""&VbCrLf&"</head>"&VbCrLf&"<body><div id=""ad_usertop""></div>"&show
		
        show=show&"<div id=""powered""><a href=""http://www.myhomestay.com.cn"" target=""_blank""><img src=""images/meiqee.com.gif"" border=""0"" alt=""Powered by meiqee.com"" /></a></div>"&VbCrLf&"<div id=""ad_userbot""></div></body>"&VbCrLf&"</html>"
        show="<title>"&log_title&"</title>"&VbCrLf&show
        show="<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />"&VbCrLf&show
        show="<meta http-equiv=""Content-Language"" content=""zh-CN"" />"&VbCrLf&show
        show="<html>"&VbCrLf&"<head>"&VbCrLf&show
		 If InStr(show, "<div id=""oblog_edit"">") Then
            show = show & "<script src=""" & blogurl & "commentedit.asp""></script>" & vbCrlf
            show = show & "<script src=""" & blogurl & "count.asp?action=code31""></script>" & vbCrlf
        End If
        If InStr(show, "<div id=""blogzhai"">") Then
            show = show & "<script src=""" & blogurl & "inc/inc_zhai.js""></script>" & vbCrlf
        End If
        show = show & "<script src=""" & blogurl & "count.asp?action=site&id=" & user_id & """></script>" & vbCrlf
        show = show & "<script src=""" & blogurl & "login.asp?action=showindexlogin_team""></script>" & vbCrlf
        show =repl_ad(show,1)
        filt_pwblog = show
    End Function
	
	public sub update_partlog(uid,lid)'lid Ϊ lastlogid ���½��ȿ��ƣ�
		dim rs1,p,lastid,i
		i=1
		lid=clng(lid)
		Set rs1 = Server.CreateObject("Adodb.RecordSet")
        rs1.open "select top 200 logid from oBlog_log where userid=" & uid & " and isdraft=0 and logid>"&lid, conn, 1, 1
		if rs1.eof then
			rs1.Close
			set rs1=nothing
			progress 100, "���������������!"
			exit sub
		end if
        While Not rs1.EOF
            p = rs1.recordcount + 1
            progress Int(i / p * 100), "����IDΪ" & rs1(0) & "������..."'ѭ��������ʾ��
            Update_log rs1(0), 0
            Update_calendar rs1(0)
            lastid=rs1(0)
			i=i+1
            rs1.movenext
        Wend
		set rs1=oblog.execute("select top 1 logid from oblog_log where userid=" & uid & " and isdraft=0 and logid>"&lastid)
		if rs1.eof then set rs1=nothing:exit sub
		rs1.Close
		set rs1=nothing
		If P_Blog_UPDATEPAUSE > 0 Then
			progress 100, P_Blog_UPDATEPAUSE&"����Զ����º�������£��벻Ҫˢ��ҳ��..."
			with response
				.Write "<script language=JavaScript>var secs = "&P_Blog_UPDATEPAUSE&";var wait = secs * 1000;"
				.write "for(i = 1; i <= secs; i++){window.setTimeout(""Update("" + i + "")"", i * 1000);}"
				.write "function Update(num){if(num != secs){printnr = (wait / 1000) - num;pstr.innerHTML=printnr+""����Զ����º�������£��벻Ҫˢ��ҳ��..."";progress.style.width=(num/secs)*100+""%"";progress.innerHTML=""ʣ��""+printnr+""��""}}"
				.write "setTimeout(""window.location='user_update.asp?action=update_alllog&lastlogid="&lastid&"'"","&int(P_Blog_UPDATEPAUSE*1000)&");</script>"
			end with
			response.Flush()
			response.End()
        End If
	end sub
	
	public sub update_usite(uid)
		dim p
		p = 12
        progress Int(1 / p * 100), "������ҳ..."
        	Update_index 0
        progress Int(2 / p * 100), "����վ����Ϣ�ļ�..."
        	Update_info uid
        progress Int(3 / p * 100), "�����������б��ļ�..."
        	Update_newblog (uid)
        progress Int(4 / p * 100), "������������..."
        	Update_newmessage uid
        progress Int(5 / p * 100), "������ҳ���·����ļ�..."
        	Update_Subject (uid)
        progress Int(6/ p * 100), "���¹���ҳ..."
        	CreateFunctionPage
        progress Int(7 / p * 100), "�������԰�..."
        	Update_message 0
        progress Int(8 / p * 100), "�������»ظ�..."
        	Update_comment uid
        progress Int(9 / p * 100), "���¹���..."
        	Update_placard uid
        progress Int(10 / p * 100), "������������..."
        	Update_links uid
        progress Int(11 / p * 100), "����BlogName..."
        	Update_blogname
        progress Int(12 / p * 100), "��վ���·������!"
	end sub
	public sub update_alllog_admin(uid)
		dim rs1,i,p
        uid = CLng(uid)
        oblog.CreateUserDir uid, 0
        userid = uid
		Set rs1 = Server.CreateObject("Adodb.RecordSet")
        rs1.open "select logid from oBlog_log where userid=" & uid & " and isdraft=0", conn, 1, 1
        While Not rs1.EOF
            p = rs1.recordcount + 1
            progress Int(i / p * 100), "����IDΪ" & rs1(0) & "������..."
            Update_log rs1(0), 0
            Update_calendar rs1(0)
            i = i + 1
            rs1.movenext
        Wend
        rs1.Close
		set rs1=nothing
		update_usite(uid)'ȫ����¸�����Ŀ��
	end sub
    
    Public Sub Update_alllog(uid,lastlogid)
        Dim p
        uid = CLng(uid)
        oblog.CreateUserDir uid, 0
        userid = uid
        update_partlog uid,lastlogid
        update_usite(uid)
    End Sub
    Public Sub Update_blogname()
        savefile user_path, "\inc\show_blogname.htm", oblog.filt_html(BlogName)
    End Sub
    Public Sub savefile(dirstr, fname, str)
        On Error Resume Next
        Dim dirstr1, divid
		if dirstr="" then
			response.Write("�û�Ŀ¼����Ϊ�գ�")
			response.end
		end if
        dirstr1 = Server.Mappath(blogdir&dirstr)
        'Response.Write "Ŀ¼2��" & dirstr1  & "<br>"
		'����תΪjs��ʽ
        If (Left(fname, 5) = "\inc\" Or Left(fname, 10) = "\calendar\") And (f_ext = "htm" Or f_ext = "html") Then
            If Left(fname, 10) = "\calendar\" Then
                divid = "calendar"
            Else
                divid = Replace(Replace(Replace(fname, "\inc\", ""), ".htm", ""), "show_", "")
            End If
            str = oblog.htm2js_div(str, divid)
        End If
		'���¼���asp��ʽ,ת��·��
		if f_ext="asp" and true_domain=0 then
			if logfilepath=0  then
				str=replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""")
			else
				if instr(fname,"index.asp") or instr(fname,"message.asp") or instr(fname,"cmd.asp") then
					str=replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""")
				else
					str=replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""../../")
				end if
			end if
			
			if logfilepath=0  then
				str=replace(str,"<!-- #include file="""&blogdir,"<!-- #include file=""../../")
			else
				if instr(fname,"index.asp") or instr(fname,"message.asp") or instr(fname,"cmd.asp") then
					str=replace(str,"<!-- #include file="""&blogdir,"<!-- #include file=""../../")
				else
					str=replace(str,"<!-- #include file="""&blogdir,"<!-- #include file=""../../../../")
				end if
			end if
		end if
		if true_domain=1 then
			if logfilepath=0  then
				str=replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""")
			else
				if instr(fname,"index."&f_ext) or instr(fname,"message."&f_ext) or instr(fname,"MyHomestay."&f_ext) then
					str=replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""")
				else
					str=replace(str,"<!-- #include file="""&user_truepath,"<!-- #include file=""../../")
				end if
			end if
		end if
        If str = "" Or IsNull(str) Then str = " "
        If objFSO.FolderExists(dirstr1) = False Then objFSO.CreateFolder (dirstr1)
        Call oblog.BuildFile(dirstr1 & Trim(fname), str)
       'If Err Then
            'err.Clear
            'Response.Write "<br />�����ļ��������������ǵ�һ��ʹ�ã���ѡ��>>����>>���·���ȫվ��"
            'Response.End
        'End If
    End Sub
    
    Public Function newcalendar(folderspec)
        'Response.Write(folderspec)
        On Error Resume Next
        Dim f, f1, fc, nName
        If objFSO.FolderExists(Server.Mappath(folderspec)) Then
            Set f = objFSO.GetFolder(Server.Mappath(folderspec))
            Set fc = f.Files
            nName = 0
            For Each f1 In fc
                If nName < CLng(Replace(f1.Name, ".htm", "")) Then nName = CLng(Replace(f1.Name, ".htm", ""))
            Next
            newcalendar = nName
          Else
            newcalendar = "0"
        End If
    End Function
    
    'str��Ϊ�˴����js�����ļ�����������ҳ�棡��
    Public Function repl_label(show, str, title, author, keyword, desc, calendar)
        On Error Resume Next
        show = Replace(show, "#ad_usercomment#", "<div id=""ad_usercomment""></div>")
        show = Replace(show, "$show_placard$", "<div id=""placard""><!-- #include file="""&user_truepath&"inc/show_placard.htm"" --></div>")
       	show = Replace(show, "$show_calendar$", "<div id=""calendar""><!-- #include file="""&user_truepath&"calendar/" & calendar & ".htm"" --></div>")
        show = Replace(show, "$show_xml$", "<div id=""xml""><a href="""&user_truepath&"rss2.xml"" target=""_blank""><img src='" & blogurl & "images/xml.gif' width='36' height='14' border='0' /></a></div>")
        show = Replace(show, "$show_subject$", "<div id=""subject""><!-- #include file="""&user_truepath&"inc/show_subject.htm"" --></div>")
        show = Replace(show, "$show_subject_l$", "<div id=""subject_l""><!-- #include file="""&user_truepath&"inc/show_subject.htm"" --></div>")
        show = Replace(show, "$show_newblog$", "<div id=""newblog""><!-- #include file="""&user_truepath&"inc/show_newblog.htm"" --></div>")
        show = Replace(show, "$show_comment$", "<div id=""comment""><!-- #include file="""&user_truepath&"inc/show_comment.htm"" --></div>")
        show = Replace(show, "$show_blogname$", "<span id=""blogname""><!-- #include file="""&user_truepath&"inc/show_blogname.htm"" --></span>")
        show = Replace(show, "$show_newmessage$", "<div id=""newmessage""><!-- #include file="""&user_truepath&"inc/show_newmessage.htm"" --></div>")
        show = Replace(show, "$show_links$", "<div id=""links""><!-- #include file="""&user_truepath&"inc/show_links.htm"" --></div><div id=""ad_userlinks""></div>")
        show = Replace(show, "$show_info$", "<div id=""info""><!-- #include file="""&user_truepath&"inc/show_info.htm"" --></div>")
        show = Replace(show, "$show_search$", "<div id=""search""><!-- #include file="""&user_truepath&"inc/show_search.htm"" --></div>")
        show = Replace(show, "$show_login$", "<div id=""ob_login""></div>")
		show = Replace(show, "$show_userlogin$", "<div id=""ob_login""></div>")
		show = Replace(show, "$show_userlogin_team$", "<div id=""ob_login_team""></div>")
		
		Call Nav_indexshow_Nav(show)'WL;Update Article Page + Update ��Ŀ Page;
		
		'show = Replace(show, "$show_login_team$", "<div id=""ob_login_team""></div>")
        'show = "<link rel=""alternate"" href="""&user_truepath&"rss2.xml"" type=""application/rss+xml"" title=""RSS"" />"&vbCrLf&"" & vbCrLf & "<script src=""" & blogurl & "inc/main.js"" type=""text/javascript""><'/script>" & vbCrLf & "</head>" & vbCrLf & "<body>" & vbCrLf & "<div id=""ad_usertop"">333</div>" & show'�ڴ���ǰ����ͷ�ļ�wl��
		show = "<link rel=""alternate"" href="""&user_truepath&"rss2.xml"" type=""application/rss+xml"" title=""RSS"" />"&vbCrLf&"" & vbCrLf & "<script src=""" & blogurl & "inc/main.js"" type=""text/javascript""></script>" & vbCrLf & "</head>" & vbCrLf & "<body>" & vbCrLf & "<div id=""ad_usertop""></div>" & show'�ڴ���ǰ����ͷ�ļ�wl��
		
        show = show & "<div id=""ad_userbot""></div>"'�ڴ���ײ�������β�ļ�wl��
        show = show & "<div id=""powered"" style=""text-align:center;""><a href=""http://www.myhomestay.com.cn"" target=""_blank""><img src="""&blogurl&"images/meiqee.com.gif"" border=""0"" alt=""Powered by myhomestay.com.cn"" /></a></div>" & vbCrLf & "</body>" & vbCrLf & "</html>"'html�����βwl��
        show = "<title>" & Replace(Replace(Replace("�ҵ�ס���� Homestay Beijing --"& title &"", "&lt;", "��"), "&gt;", "��"), "&nbsp;", "��") & "</title>" & vbCrLf & show
        show = "<meta name=""description"" content=""" & oblog.filt_html(desc) & """ />" & vbCrLf & show
        show = "<meta name=""keyword"" content=""" & oblog.filt_html(keyword) & """ />" & vbCrLf & show
        show = "<meta name=""author"" content=""" & oblog.filt_html(author) & """ />" & vbCrLf & show
        show = "<meta name=""generator"" content=""myhomestay.com.cn"" />" & vbCrLf & show
        show = "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" & vbCrLf & show
        show = "<meta http-equiv=""Content-Language"" content=""zh-CN"" />" & vbCrLf & show
       ' show = "<html>" & vbCrLf & "<head>" & vbCrLf & show

		show = "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbcrlf & "<head>" & vbCrLf & show
		show = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbCrLf & show
        'html�ļ����������ļ���Ϊjs����
        If f_ext = "htm" Or f_ext = "html" Then
            show = filt_include(show)
            show = show & "<script src=""" & user_truepath&"inc/show_subject.htm""></script>" & vbCrlf
           '''WL show = show & "<script src=""" & user_truepath&"inc/show_placard.htm"">< /script>" & vbCrlf
           '''WL show = show & "<script src=""" & user_truepath&"calendar/" & calendar & ".htm"">< /script>" & vbCrlf
            show = show & "<script src=""" & user_truepath&"inc/show_search.htm""></script>" & vbCrlf
            show = show & "<script src=""" & user_truepath&"inc/show_newblog.htm""></script>" & vbCrlf
            show = show & "<script src=""" & user_truepath&"inc/show_comment.htm""></script>" & vbCrlf
            show = show & "<script src=""" & user_truepath&"inc/show_newmessage.htm""></script>" & vbCrlf
            show = show & "<script src=""" & user_truepath&"inc/show_info.htm""></script>" & vbCrlf
            show = show & "<script src=""" & user_truepath&"inc/show_links.htm""></script>" & vbCrlf
            'show = show & "<script src=""" & user_truepath&"inc/show_xml.htm""><script>" & vbCrlf
            show = show & "<script src=""" & user_truepath&"inc/show_blogname.htm""></script>" & vbCrlf
        End If
        '����Ϊshtml��asp��html����js����
        If InStr(show, "<div id=""oblog_edit"">") Then
            show = show & "<script src=""" & blogurl & "commentedit.asp""></script>" & vbCrlf
            show = show & "<script src=""" & blogurl & "count.asp?action=code31""></script>" & vbCrlf
        End If
        If InStr(show, "<div id=""blogzhai"">") Then
            show = show & "<script src=""" & blogurl & "inc/inc_zhai.js""></script>" & vbCrlf
        End If
        show = show & str
        show = show & "<script src=""" & blogurl & "count.asp?action=site&id=" & user_id & """></script>" & vbCrlf
        show = show & "<script src=""" & blogurl & "login.asp?action=showindexlogin""></script>" & vbCrlf
		show = show & "<script src=""" & blogurl & "login.asp?action=showindexlogin_team""></script>" & vbCrlf
		show=repl_ad(show,0)
        repl_label=filtskinpath(show)
    End Function
	
	'�滻������
	public function repl_ad(str,t)
		dim show
		show=str
		if user_level=8 then
			show = Replace(show, "<div id=""ad_usertop""></div>", "")
			show = Replace(show, "<div id=""ad_userbot""></div>", "")
			show = Replace(show, "<div id=""ad_userlinks""></div>", "")
			show = Replace(show, "<div id=""ad_usercomment""></div>", "")
			repl_ad=show
			exit function
		end if
		if t=0 then '��̬ҳ�棻
			if (f_ext="shtml" or f_ext="asp") and true_domain=0 then
				show = Replace(show, "<div id=""ad_usertop""></div>", "<!-- #include file="""&blogurl&"ad/ad_usertop.htm"" -->")
				show = Replace(show, "<div id=""ad_userbot""></div>", "<!-- #include file="""&blogurl&"ad/ad_userbot.htm"" -->")
				show = Replace(show, "<div id=""ad_userlinks""></div>", "<!-- #include file="""&blogurl&"ad/ad_userlinks.htm"" -->")
				show = Replace(show, "<div id=""ad_usercomment""></div>", "<!-- #include file="""&blogurl&"ad/ad_usercomment.htm"" -->")
				
			elseif f_ext="html" or f_ext="htm" or true_domain=1 then
				show = Replace(show, "<div id=""ad_usertop""></div>", "<script src=""" & blogurl & "ad/ad_usertopjs.htm""></script>")
				show = Replace(show, "<div id=""ad_userbot""></div>", "<script src=""" & blogurl & "ad/ad_userbotjs.htm""></script>")
				show = Replace(show, "<div id=""ad_userlinks""></div>", "<script src=""" & blogurl & "ad/ad_userlinksjs.htm""></script>")
				show = Replace(show, "<div id=""ad_usercomment""></div>", "<script src=""" & blogurl & "ad/ad_usercommentjs.htm""></script>")
			end if
		else
			'��̬����ֱ���滻�ɰ����ļ��е����ݣ�
			show = Replace(show, "<div id=""ad_usertop""></div>",oblog.readfile("ad","ad_usertop.htm"))
			show = Replace(show, "<div id=""ad_userbot""></div>",oblog.readfile("ad","ad_userbot.htm"))
			show = Replace(show, "<div id=""ad_userlinks""></div>", oblog.readfile("ad","ad_userlinks.htm"))
			show = Replace(show, "<div id=""ad_usercomment""></div>",oblog.readfile("ad","ad_usercomment.htm"))
		end if
		repl_ad=show
	end function
    
    Public Sub progress_init1()
        Response.Write ("<table width=""400"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"">")
        Response.Write ("<tr><td style=""text-align:center;""><span id=txt1 Name=txt1 ></span></td></tr>")
        Response.Write ("<tr><td><img src=""images/bar.gif"" width=0 height=10 id=img1 Name=img1 align=absmiddle></td></tr>")
        Response.Write ("<tr><td><span id=txt2 Name=txt2 ></span></td></tr></table>")
    End Sub
    
    Public Sub progress_init()
        Response.Write ("ϵͳ��Ϣ��<span id=""pstr""></span><div class=""progress1""><div class=""progress2"" id=""progress""></div></div>")
    End Sub
    
    Public Sub progress(num, str)
        Response.Write "<script>progress.style.width=""" & num & "%"";progress.innerHTML=""" & num & "%"";pstr.innerHTML=""" & str & """;</script>" & vbCrLf
        Response.Flush'�����ٲ�����256�ֽڵ����ݺ����̽���������ݴ����ͻ�����ʾ�ˣ�
    End Sub
    
    '����һ�������ܽ��棬�书�ܱ�ǩ����$show_log$
    '����ʱ��
    '����/�޸�����/�����ظ�/��������/ģ�����/��վ����
    Public Sub CreateFunctionPage()
    	Page="cmd"
        'On Error Resume Next
        Dim strPage,str
		str=user_skin_main_Nav
       	strPage = repl_label(str, "", BlogName, user_userName & "," & user_nickName, "", "", "")     
		strPage=Replace(strPage,"$show_log$","<div id=oblog_usercontent>"&oblog.setup(82,0)&"</div>") 
        strPage = strPage & "<script>document.write('<script language=""javascript"" src="""&blogurl&"pagecmd.asp?'+getpara()+'""><\/script>');</script>"
		if f_ext="shtml" or f_ext="asp" then
			Dim objRegExp
			Set objRegExp = New Regexp  
			objRegExp.IgnoreCase = True
			objRegExp.Global = True
			objRegExp.Pattern = "<div id=""calendar"">.*</div>"
			strPage = objRegExp.Replace(strPage, "<div id=""calendar"">"&oblog.setup(82,0)&"</div>")
		  	Set objRegExp = Nothing
		end if
        Savefile user_path, "\MyHomestay."&f_ext, strPage
        strPage=""
        Page=""
    End Sub
    '������̬��ҳҳ��ķ�ҳ������
    Private Function CreateStaticPageBar(byval lngAll,byval intPerPage,byval intType)
    	Dim strPageBar,strFileName,strUnit
    	strFileName=user_truepath&"MyHomestay."&f_ext&"?uid=" & user_id& "&do="'���ӵ�ַ����ҳѰcmd.html��
    	Select Case intType
    		Case 0
    			strFileName=strFileName & "index"
    			strUnit="ƪ����"
    		Case 1
    			strFileName=strFileName & "message"
    			strUnit="ƪ����"
    	End Select 
    	CurrentPage=1 	
    	CreateStaticPageBar=oBlog.ShowPage(strFileName, lngAll, intPerPage, false, true, strUnit)
    End Function

End Class
%>

<%

Sub Nav_indexshow_Nav(showTemp)
'    If InStr(show, "$show_placard$") > 0 Then
'        show = Replace(show, "$show_placard$", show_placard())
'    End If
    
    If InStr(showTemp, "$show_friends$") > 0 Then
    showTemp = Replace(showTemp, "$show_friends$", show_friends())
    End If
    
'    If InStr(show, "$show_count$") > 0 Then
'        show = Replace(show, "$show_count$", show_count())
'    End If
    
'    If InStr(show, "$show_userlogin$") > 0 Then
'        show = Replace(show, "$show_userlogin$", show_userlogin())
'    End If
'	If InStr(show, "$show_userlogin_team$") > 0 Then
'        show = Replace(show, "$show_userlogin_team$", show_userlogin_team())
'    End If
    
'    If InStr(show, "$show_xml$") > 0 Then
'        show = Replace(show, "$show_xml$", show_sysxml())
'    End If
    
''    If InStr(showTemp, "$show_blogstar$") > 0 Then
''        showTemp = Replace(showTemp, "$show_blogstar$", show_blogstar())
''    End If

	If InStr(showTemp, "$GetSysClassesSearchHome_start$") > 0 Then'��ȡ���ڵ���(ʡ/��)����ѧϰ�������ࡪ�����ݽṹ������ͥ��Ų�ѯ��
        showTemp = Replace(showTemp, "$GetSysClassesSearchHome_start$", GetSysClassesSearchHome_start())
    End If
	
	
    Call Nav_runsub_Nav("$show_newblogger",showTemp)
'    Call runsub("$show_comment")
'    Call runsub("$show_subject")
    Call Nav_runsub_Nav("$show_blogupdate",showTemp)
    Call Nav_runsub_Nav("$show_bestblog",showTemp)
    Call Nav_runsub_Nav("$show_bloger",showTemp)
    Call Nav_runsub_Nav("$show_class",showTemp)
'    Call runsub("$show_log")
'    Call runsub("$show_userlog")
'    Call runsub("$show_search")
    Call Nav_runsub_Nav("$show_cityblogger",showTemp)
    Call Nav_runsub_Nav("$show_newphoto",showTemp)
    Call Nav_runsub_Nav("$show_blogstar2",showTemp)
	
	Nav_indexshow_Nav = showTemp
End Sub

Sub Nav_runsub_Nav(label,showTemp)
    On Error Resume Next
    Dim tmp1, tmp2, i
    Dim tmpstr, para
    tmp2 = 1
    While InStr(tmp2, showTemp, label) > 0
        tmp1 = InStr(tmp2, showTemp, label)
        tmp2 = InStr(tmp1 + 1, showTemp, "$")
        tmpstr = Mid(showTemp, tmp1, tmp2 - tmp1)
        tmpstr = Replace(tmpstr, "(", "")
        tmpstr = Replace(tmpstr, ")", "")
        tmpstr = Trim(Replace(tmpstr, label, ""))
        para = Split(tmpstr, ",")
        Select Case label
'        Case "$show_log"
'            show=replace(show,label&"("&tmpstr&")$",show_log(para(0),para(1),para(2),para(3),para(4),para(5),para(6),para(7),para(8)))
'            If Err Then
'                Response.Write "<br/>$show_log$��ǩ�д����������"
'                Response.End()
'            End If
'        Case "$show_userlog"
'            show=replace(show,label&"("&tmpstr&")$",show_userlog(para(0),para(1),para(2),para(3),para(4),para(5)))
'            If Err Then
'            	Response.Write Err.Description
'                Response.Write "<br/>$show_userlog$��ǩ�д����������"
'                Response.End()
'            End If
'        Case "$show_comment"
'            show=replace(show,label&"("&tmpstr&")$",show_comment(para(0),para(1)))
'            If Err Then
'                Response.Write "<br/>$show_comment$��ǩ�д����������"
'                Response.End()
'            End If
'        Case "$show_subject"
'            show=replace(show,label&"("&tmpstr&")$",show_subject(para(0)))
'            If Err Then
'                Response.Write "<br/>$show_subject$��ǩ�д����������"
'                Response.End()
'            End If
        Case "$show_blogupdate"
            showTemp=replace(showTemp,label&"("&tmpstr&")$",show_blogupdate(para(0)))
            If Err Then
                Response.Write "<br/>$show_blogupdate$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_newblogger"
            showTemp=replace(showTemp,label&"("&tmpstr&")$",show_newblogger(para(0)))
            If Err Then
                Response.Write "<br/>$show_newblogger$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_bestblog"
            showTemp=replace(showTemp,label&"("&tmpstr&")$",show_bestblog(para(0)))
            If Err Then
                Response.Write "<br/>$show_bestblog$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_bloger"
            showTemp=replace(showTemp,label&"("&tmpstr&")$",show_bloger(para(0)))
            If Err Then
                Response.Write "<br/>$show_bloger$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_class"
            showTemp=replace(showTemp,label&"("&tmpstr&")$",show_class(para(0)))
            If Err Then
                Response.Write "<br/>$show_class$��ǩ�д����������"
                Response.End()
            End If
'        Case "$show_search"
'            show=replace(show,label&"("&tmpstr&")$",show_search(para(0)))
'            If Err Then
'                Response.Write "<br/>$show_search$��ǩ�д����������"
'                Response.End()
'            End If
        Case "$show_cityblogger"
            showTemp=replace(showTemp,label&"("&tmpstr&")$",show_cityblogger(para(0)))
            If Err Then
                Response.Write "<br/>$show_cityblogger$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_newphoto"
            showTemp=replace(showTemp,label&"("&tmpstr&")$",show_newphoto(para(0),para(1),para(2),para(3)))
            If Err Then
                Response.Write "<br/>$show_newphoto$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_blogstar2"
            showTemp=Replace(showTemp,label&"("&tmpstr&")$",show_blogstar2(para(0),para(1),para(2),para(3)))
            If Err Then
                Response.Write "<br/>$show_blogstar2$��ǩ�д����������"
                Response.End()
            End If
        End Select
    Wend
	Nav_runsub_Nav = showTemp
End Sub


Function show_cityblogger(i)
    show_cityblogger = "<form name=""oblogform"" action=""/HomestaySearchcityuser.asp"" target=""_blank"">" & oblog.type_city("", "") & " <input type='submit' value='����'></form>"
    If i = 1 Then show_cityblogger = Replace(show_cityblogger, "<select name='city'", "<br /><select name='city'")
End Function



Function show_class(m)
    Dim rs
    'show_class="<a href=index.asp>��ҳ("&blogcount&")</a><br>"
    Dim i, brstr
    show_class = ""
    m = Int(m)
    Set rs = oblog.execute("select id,classname,userid_Nav,EName,root_link  from oblog_logclass Where idtype=0 And isDisplay=1 order by RootID,OrderID")
    If m = 0 Then'������ʾ��
        While Not rs.EOF
            show_class=show_class&"<a href=HomestayNav.asp?classid="&rs(0)&">"&rs(1)&"</a><br/>"
            rs.MoveNext
        Wend
    Else'm=1��m>1��ÿ�ж��ٸ���ʾ�������;
        i = 0
        While Not rs.EOF
            i = i + 1
            If i = Int(m) Then
                brstr = "<br/>"
                i = 0'�������㣻
            Else
                brstr = ""
            End If
            'show_class=show_class&"<a href=HomestayNav.asp?classid="&rs(0)&">"&rs(1)&"</a> "&brstr
				If rs("userid_Nav")>0 And rs("root_link")="" Then'�������޸�֮��WL;(���û�й̶�����)
					Dim sqlTmpNav,rsTmpNav,uname,domain,hideurl,reurl,uid
					sqlTmpNav="select user_dir,userid,hideurl,user_folder,user_domain,user_domainroot"&str_domain&" from oblog_user where userid="&rs("userid_Nav")
					set rsTmpNav=oblog.execute(sqlTmpNav)
					if not rsTmpNav.eof then
						hideurl=rsTmpNav(2)
						if true_domain=1 then
							if rsTmpNav("custom_domain")="" or isnull(rsTmpNav("custom_domain")) then
								reurl="http://"&rsTmpNav("user_domain")&"."&rsTmpNav("user_domainroot")
							else
								reurl="http://"&rsTmpNav("custom_domain")
							end if
						else
							reurl=blogdir&rsTmpNav(0)&"/"&rsTmpNav(3)
						end if
						
							if oblog.setup(81,0)=1 then
								'''reurl=reurl&"/index."&f_ext
								reurl=reurl&"/"'[2007-7-27]WL;
							else
								reurl=reurl
							end if
						
						set rsTmpNav=nothing
						
					else
						set rsTmpNav=nothing
						oblog.adderrstr("�����޴���Ŀ�ռ�!")
						oblog.showerr
					end if
					
					'show_class=show_class&"<li><a href="&reurl&" target='_blank'>"&rs(1)&"</a></li>"&brstr
					show_class=show_class&"<li><a href="& reurl &">"& rs(1)
					If rs("EName")<>"" Then show_class=show_class&"<br />"& rs("EName")
					show_class=show_class&"</a></li>"& brstr
					
				ElseIf rs("root_link")<>"" And Len(rs("root_link"))>6 Then'�������޸�֮��WL;(����й̶�����...����ʹ�ù̶�����)
					show_class=show_class&"<li><a href="& rs("root_link") &">"&rs(1)					
					If rs("EName")<>"" Then show_class=show_class&"<br />"& rs("EName")
					show_class=show_class&"</a></li>"& brstr
				Else
					show_class=show_class&"<li><a href=HomestayNav.asp?classid="&rs(0)&">"&rs(1)					
					If rs("EName")<>"" Then show_class=show_class&"<br />"& rs("EName")
					show_class=show_class&"</a></li>"& brstr
				End If
			rs.MoveNext
        Wend
    End If
    Set rs = Nothing
	
	Dim dirstr
	dirstr= Server.MapPath("menu/")
	'response.Write ">"
	'ר��user���ɣ�Call oblog.BuildFile(dirstr&"\incIndexNav.htm",show_class)'user��̨���õ���Ŀ�����ļ�WL;
	
End Function

Function show_friends()
    If IsNull(oblog.setup(19, 0)) Then
        show_friends = " "
    Else
        show_friends = oblog.setup(19, 0)
    End If
End Function


'��ȡ���ڵ���(ʡ/��)����ѧϰ�������ࡪ�����ݽṹ������ͥ��Ų�ѯ��
Function GetSysClassesSearchHome_start()
	Dim strType_city
	Dim str_usertype
	Dim rst,sReturn
	
	Dim Request_province,Request_city
	Request_province= Request("province")
	Request_city= Request("city")
	
	Dim Request_usertype
	Request_usertype= Cint(Request("usertype"))
	
	Dim Request_houseframe
	Request_houseframe= Trim(Request("houseframe"))
	
	strType_city =strType_city & "<Div id='line01' style='margin:0px;padding-top:3px;'><p style='margin:0px;padding:0px; margin-top:5px; margin-bottom:5px;'>&nbsp;<font color=black style='font-size:12px;'>���ڵ���(ʡ/��)</font></p>" &oblog.type_citySelectSubmit(Request_province,Request_city )'����ѡ�������б�
	
	
	str_usertype="<select name=usertype id=usertype onchange='submit();'>"
	str_usertype=str_usertype&oblog.show_class("user",Request_usertype,0)
    str_usertype=str_usertype&"</select>"	
	str_usertype = "<br /><p style='margin:0px;padding:0px; margin-top:10px; margin-bottom:5px;'>&nbsp;<font color=black style='font-size:12px;'>�����ֲ�ѯ<br /></font></p>" &str_usertype&"</Div>"'����ѡ�������б�
	
'	Set rst=conn.Execute("Select * From oblog_logclass Where idtype=1")
'	If rst.Eof Then
'		sReturn=""
'	Else
'		Do While Not rst.Eof
'			sReturn= sReturn & "<option value="&rst("id")&">" & rst("classname") & "</option>" & VBCRLF
'			rst.Movenext
'		Loop
'		sReturn = "<option value=0>���з���</option>" & VBCRLF & sReturn
		sReturn="<form name=selectsubmitform action='/HomestayFamilyChina.asp' method=get><tr><td clospan=2>" & strType_city & str_usertype & "<Div id='line02' style='margin:0px;padding-top:6px;'><p style='margin:0px;padding:0px; margin-bottom:5px;'>&nbsp;<font color=black style='font-size:12px;'>���ݽṹ</font></p><select name='houseframe' id='houseframe1' onchange='submit();' title='��ѡ�����ķ��ݽṹ'>"&_
                            "<option value='' selected>��</option>"&_
							"<option value='1��'"
							If Instr(Request_houseframe,"1��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">1��</option>"&_
                            "<option value='2��'"
							If Instr(Request_houseframe,"2��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">2��</option>"&_
                            "<option value='3��'"
							If Instr(Request_houseframe,"3��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">3��</option>"&_
                            "<option value='4��'"
							If Instr(Request_houseframe,"4��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">4��</option>"&_
                            "<option value='����'"
							If Instr(Request_houseframe,"����")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">����</option>"&_
                            "<option value='����'"
							If Instr(Request_houseframe,"����")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">����</option>"&_
                          "</select>&nbsp;<select name='houseframe' id='houseframe2' onchange='submit();' title='��ѡ�����ķ��ݽṹ'>"&_
						  "<option value='' selected>��</option>"&_
                            "<option value='1��'"
							If Instr(Request_houseframe,"1��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">1��</option>"&_
                            "<option value='2��'"
							If Instr(Request_houseframe,"2��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">2��</option>"&_
                            "<option value='3��'"
							If Instr(Request_houseframe,"3��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">3��</option>"&_
                            "<option value='4��'"
							If Instr(Request_houseframe,"4��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">4��</option>"&_
                            "<option value='5��'"
							If Instr(Request_houseframe,"5��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">5��</option>"&_
                            "<option value='6��'"
							If Instr(Request_houseframe,"6��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">6��</option>"&_
                          "</select>&nbsp;<select name='houseframe' id='houseframe3' onchange='submit();' title='��ѡ�����ķ��ݽṹ'>"&_
						  "<option value='' selected>��</option>"&_
                            "<option value='1��'"
							If Instr(Request_houseframe,"1��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">1��</option>"&_
                            "<option value='2��'"
							If Instr(Request_houseframe,"2��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">2��</option>"&_
                            "<option value='3��'"
							If Instr(Request_houseframe,"3��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">3��</option>"&_
                            "<option value='4��'"
							If Instr(Request_houseframe,"4��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">4��</option>"&_
                            "<option value='5��'"
							If Instr(Request_houseframe,"5��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">5��</option>"&_
                            "<option value='6��'"
							If Instr(Request_houseframe,"6��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">6��</option>"&_
                          "</select></Div> </tr></form>"
						  
						  
	sReturn=sReturn& "<form name=selectsubmitform2 action='/HomestayFamilyChina.asp' method=post><tr><td clospan=2>"
		sReturn=sReturn& "<Div style='font-size:12px;margin:13px;margin-top:28px;'><font color='#BE0000' style='font-weight:bold;'>ֱ�Ӳ�ѯ��ͥ���</font>"
			sReturn=sReturn& "<input type='text' name='SearchUserId' value='' size='15' style='height:15px;font-weight:normal;' />&nbsp;&nbsp;"
			sReturn=sReturn& "<input type='submit' name='submit' value='��ѯ' />"
			sReturn=sReturn& "<input type='hidden' name='UserSearch' value='10' />"
		sReturn=sReturn& "</Div>"
	sReturn=sReturn& "<td></tr></form>"
'	End If
'	rst.Close
'	Set rst=Nothing
	GetSysClassesSearchHome_start = sReturn
End Function

%>