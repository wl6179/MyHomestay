<%
Dim show
Dim totalPut, CurrentPage, TotalPages, MaxPerPage, strFileName
show=show& "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
show=show& "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbcrlf
show=show& "<head>" & vbcrlf
show=show& "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"& vbcrlf
'show=show& "<META name='verify-v1' content='4AFPtwVeCSHTeNJ4hjkCxfSl1bh8NtWS4oo409zoOMs=' />"& vbcrlf
show=show& "<meta name=""generator"" content=""oblog"">"& vbcrlf
show=show& "<meta name='keywords' content='"&oblog.setup(5,0)&"'>"& vbcrlf
show=show& "<link rel='alternate' href='rss2.asp' type='application/rss+xml' title='RSS' >"& vbcrlf
show=show& "<title>"&oblog.setup(2,0)&"</title>"& vbcrlf
show=show& "</head>"& vbcrlf
show=show&"<script>function chkdiv(divid){var chkid=document.getElementById(divid);if(chkid != null){return true; }else {return false; }}</script>"

Function show_userlogin()
    show_userlogin = "<div id=""ob_login""></div>"
End Function

Sub indexshow()
    If InStr(show, "$show_placard$") > 0 Then
        show = Replace(show, "$show_placard$", show_placard())
    End If
    
    If InStr(show, "$show_friends$") > 0 Then
    show = Replace(show, "$show_friends$", show_friends())
    End If
    
    If InStr(show, "$show_count$") > 0 Then
        show = Replace(show, "$show_count$", show_count())
    End If
    
    If InStr(show, "$show_userlogin$") > 0 Then
        show = Replace(show, "$show_userlogin$", show_userlogin())
    End If
    
    If InStr(show, "$show_xml$") > 0 Then
        show = Replace(show, "$show_xml$", show_sysxml())
    End If
    
    If InStr(show, "$show_blogstar$") > 0 Then
        show = Replace(show, "$show_blogstar$", show_blogstar())
    End If
    
    If InStr(show, "$show_cloudtags$") > 0 Then
        show = Replace(show, "$show_cloudtags$", Tags_SystemTags("1"))
    End If
    
    Call runsub("$show_newblogger")
    Call runsub("$show_comment")
    Call runsub("$show_subject")
    Call runsub("$show_blogupdate")
    Call runsub("$show_bestblog")
    Call runsub("$show_bloger")
    Call runsub("$show_class")
    Call runsub("$show_log")
	
	'Begin������ӵģ���ʾ Top��ע����
	'����ӵ��±�� show_toplog����ʾ Top��ע�������б���91health����ʽ���ͽ�����Ѷ�ͽ������̸ͳһ
	
	Call runsub("$show_toplog")
	
	'End��***********************
	
	
		
    Call runsub("$show_userlog")
    Call runsub("$show_search")
    Call runsub("$show_cityblogger")

	'����ӵģ�����ҳ��ʾͬ���û�
	Call runsub("$show_cityblogger91")	
	'If InStr(show, "$show_cityblogger91$") > 0 Then
        'show = Replace(show, "$show_cityblogger91$", show_cityblogger91())
	'	Call runsub("$show_cityblogger91")	
    'End If
    
	Call runsub("$show_newphoto")
    Call runsub("$show_blogstar2")
    'Call runsub("$show_hotTags")
End Sub

Sub sysshow()
	if Application(oblog.cache_name&"list_update")=false and application(oblog.cache_name&"list")<>"" then
	    show=application(oblog.cache_name&"list")
	Else
	    Dim rstmp
	    Set rstmp = oblog.execute("select skinshowlog from oblog_sysskin where isdefault=1")
	    show=show&rstmp(0)
	    Set rstmp = Nothing
	    '��ģ��ȡ������ѡ��
	    show=Replace(show,"show_cityblogger(0)$","")
	    show=Replace(show,"show_cityblogger(1)$","")
		
		'����ӵ������������д���
		show=Replace(show,"show_cityblogger91(0)$","")
	    show=Replace(show,"show_cityblogger91(1)$","")
		
	    Call indexshow
	    Application.Lock
	    application(oblog.cache_name&"list_update")=false
	    application(oblog.cache_name&"list")=show
	    Application.unLock
	End If
End Sub

Sub runsub(label)
    On Error Resume Next
    Dim tmp1, tmp2, i
    Dim tmpstr, para
    tmp2 = 1
    While InStr(tmp2, show, label) > 0
        tmp1 = InStr(tmp2, show, label)
        tmp2 = InStr(tmp1 + 1, show, "$")
        tmpstr = Mid(show, tmp1, tmp2 - tmp1)
        tmpstr = Replace(tmpstr, "(", "")
        tmpstr = Replace(tmpstr, ")", "")
        tmpstr = Trim(Replace(tmpstr, label, ""))
        para = Split(tmpstr, ",")
        Select Case label
        Case "$show_log"
            show=replace(show,label&"("&tmpstr&")$",show_log(para(0),para(1),para(2),para(3),para(4),para(5),para(6),para(7),para(8)))
            If Err Then
                Response.Write "<br/>$show_log$��ǩ�д����������"
                Response.End()
            End If
		
		'*** Begin: ����ӵģ��滻 Top��ע���� ��ǩ
		'����1��Ҫ��ʾ����������
		'����2��Ҫ��ʾ�����±����ַ����� 
		Case "$show_toplog"
            show=replace(show,label&"("&tmpstr&")$",show_toplog(para(0),para(1)))
            If Err Then
                Response.Write "<br/>$show_toplog$��ǩ�д����������"
                Response.End()
            End If
		
		'*** End *****
		
        Case "$show_userlog"
            show=replace(show,label&"("&tmpstr&")$",show_userlog(para(0),para(1),para(2),para(3),para(4),para(5)))
            If Err Then
            	Response.Write Err.Description
                Response.Write "<br/>$show_userlog$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_comment"
            show=replace(show,label&"("&tmpstr&")$",show_comment(para(0),para(1)))
            If Err Then
                Response.Write "<br/>$show_comment$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_subject"
            show=replace(show,label&"("&tmpstr&")$",show_subject(para(0)))
            If Err Then
                Response.Write "<br/>$show_subject$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_blogupdate"
            show=replace(show,label&"("&tmpstr&")$",show_blogupdate(para(0)))
            If Err Then
                Response.Write "<br/>$show_blogupdate$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_newblogger"
            show=replace(show,label&"("&tmpstr&")$",show_newblogger(para(0)))
            If Err Then
                Response.Write "<br/>$show_newblogger$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_bestblog"
            show=replace(show,label&"("&tmpstr&")$",show_bestblog(para(0)))
            If Err Then
                Response.Write "<br/>$show_bestblog$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_bloger"
            show=replace(show,label&"("&tmpstr&")$",show_bloger(para(0)))
            If Err Then
                Response.Write "<br/>$show_bloger$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_class"
            show=replace(show,label&"("&tmpstr&")$",show_class(para(0)))
            If Err Then
                Response.Write "<br/>$show_class$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_search"
            show=replace(show,label&"("&tmpstr&")$",show_search(para(0)))
            If Err Then
                Response.Write "<br/>$show_search$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_cityblogger"
            show=replace(show,label&"("&tmpstr&")$",show_cityblogger(para(0)))
            If Err Then
                Response.Write "<br/>$show_cityblogger$��ǩ�д����������"
                Response.End()
            End If
		'����ӵ�һ��Case��䣺��ʾ����ҳ��ͬ���û�
		Case "$show_cityblogger91"
            show=replace(show,label&"("&tmpstr&")$",show_cityblogger91(para(0)))
            If Err Then
                Response.Write "<br/>$show_cityblogger91$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_newphoto"
            show=replace(show,label&"("&tmpstr&")$",show_newphoto(para(0),para(1),para(2),para(3)))
            If Err Then
                Response.Write "<br/>$show_newphoto$��ǩ�д����������"
                Response.End()
            End If
        Case "$show_blogstar2"
            show=Replace(show,label&"("&tmpstr&")$",show_blogstar2(para(0),para(1),para(2),para(3)))
            If Err Then
                Response.Write "<br/>$show_blogstar2$��ǩ�д����������"
                Response.End()
            End If
        End Select
    Wend
End Sub

Function show_log(n, l, order, action, sdate, classid, classname, subjectname, info)
    show_log = ""
    Dim rs, msql, ordersql, actionsql, classsql, rstmp, i, postname, posttime, userurl, ustr
    i = 0
    Select Case order
    Case 1
       ordersql = " order by logid desc"
       'ordersql = " order by addtime desc"
    Case 2
        ordersql = " order by iis desc"
    Case 3
        ordersql = " order by commentnum desc"
    End Select
    
    Select Case action
    Case 1
        actionsql = ""
    Case 2
        actionsql = " and isbest=1"
    End Select
	
	If classid = 0 Then
        classsql = ""
    Else
        set rs=oblog.execute("select id from oblog_logclass where parentpath like '"&classid&",%' OR parentpath like '%,"&classid&"' OR parentpath like '%,"&classid&",%'")
        While Not rs.EOF
            ustr=ustr&","&rs(0)
            rs.MoveNext
        Wend
        ustr=classid&ustr
        classsql=" and classid in ("&ustr&")"
        'classsql=" and classid="&clng(classid)
    End If
    msql="select top "&n&" topic,logfile,addtime,commentnum,iis,logid,classid,subjectid,author,blogname,nickname,user_dir,user_domain,user_domainroot,user_dir,oblog_user.userid,user_folder from oblog_log,oblog_user where oblog_user.userid=oblog_log.userid and ishide<>1 and isdraft=0 and passcheck=1 and user_public<>1 "
    If is_sqldata Then
        msql=msql&" and datediff(d,oblog_log.truetime,getdate())<"&cint(sdate)&" and oblog_log.blog_password=0 And (ispassword='' Or ispassword is null)"
    Else
        msql=msql&" and datediff('d',oblog_log.truetime,now())<"&cint(sdate)&" and oblog_log.blog_password=0 And (ispassword='' Or ispassword is null)"
    End If
    msql=msql&actionsql&classsql
    msql=msql&ordersql
    Set rs = oblog.execute(msql)
    '��<a href="#" target="_blank" class="a_k">����ͳ�ƾֱ�����ͳ�ƾֱ�����ͳ�ƾֱ�����Һ̬��</a>��С�¶���<br>
    Do While Not rs.EOF
        show_log=show_log&""
        If rs("nickname") <> "" Then
            postname = oblog.filt_html(rs("nickname"))
        Else
            postname = oblog.filt_html(rs(8))
        End If
        posttime = rs(2)
        If classname = 1 Then
            Set rstmp = oblog.execute("select id,classname from oblog_logclass where id=" & rs(6))
            If Not rstmp.EOF Then
                show_log=show_log&"<a href=list.asp?classid="&rstmp(0)&" target=_blank  class='a_k' style='color:#009900'>["&rstmp(1)&"]</a>"
				'show_log=show_log&"��<a href=list.asp?classid="&rstmp(0)&" target=_blank  class='a_k'>��"&rstmp(1)&"��</a>"
			End If
		Else
			show_log=show_log&"<img src='images/blog/left-1-3.gif' vspace='7' align='absmiddle' /> "
			'show_log=show_log&"�� "
        End If
        If subjectname = 1 Then
            Set rstmp = oblog.execute("select subjectid,subjectname from oblog_subject where subjectid=" & rs(7))
            If Not rstmp.EOF Then
                show_log=show_log&"��<a href=blog.asp?name="&rs("author")&"&subjectid="&rstmp(0)&" target=_blank class='a_k'>["&oblog.filt_html(rstmp(1))&"]</a>"
            End If
        End If
        Dim topic
        If rs(0) <> "" Then
            topic = Replace(rs(0), "'", "")
            If topic <> "" Then
                If oblog.strLength(topic) > Int(l) Then
                    topic = oblog.InterceptStr(topic, Int(l) - 3) & "..."
                End If
            End If
        End If
        show_log=show_log&"<a href="""&rs(1)&""" title="""&oblog.filt_html(rs(0))&""" target=_blank class='a_k'>"&oblog.filt_html(topic)&"</a>"
        If oblog.setup(12, 0) = 1 Then
            userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
        Else
            userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
        End If
        Select Case CInt(info)
        Case 1
            show_log=show_log&"(<a href="&userurl&" target=_blank class='a_k'>"&postname&"</a>,"&formatdatetime(posttime,0)&"��"
        Case 2
            show_log=show_log&"��"&posttime&"��"
        Case 3
            show_log=show_log&"��<a href="&userurl&" target=_blank class='a_k'>"&postname&"</a>��"
        Case 4
            show_log=show_log&"��<a href="&userurl&" target=_blank class='a_k'>"&postname&"</a>,"&rs(4)&"��"
        Case 5
            show_log=show_log&"("&rs(4)&")"
        Case 6
            show_log=show_log&"��<a href="&userurl&" target=_blank class='a_k'>"&postname&"</a>,"&formatdatetime(posttime,1)&"��"
        Case 7
            show_log=show_log&"��"&formatdatetime(posttime,1)&"��"
        Case 8
            show_log=show_log&"��"&rs(3)&"��"
        Case 9
            show_log=show_log&"��"&oblog.filt_html(rs("blogname"))&"��"
        Case Else
        End Select
		
        '��Ϊ���һ����¼����
		If i < Int(n) Then  show_log=show_log&" <br>"
		
        rs.MoveNext
        i = i + 1
        If i >= Int(n) Then Exit Do
    Loop
    Set rs = Nothing
    Set rstmp = Nothing
End Function


'*** Begin: ����ӵĺ�������ʾ�� Top��ע������
'��ʾ Top��ע �����£��������
'���� n����ʾ��������
'���� l����ʾ�����±��ⳤ��
'��ʾ����ʽ������ �� �����
Function show_toplog(n, l)

    show_toplog = ""
	
    Dim rs, msql, postname, userurl, i
	
	i = 1	'�����ж�����ʾ�����л���ż����
 
	'SQL��䣺���ݵ����(iis)��������91health�е��û�����
	msql = "Select Top " & n &" topic,logfile,addtime,commentnum,iis,logid,classid,subjectid,author,blogname,nickname, " & _
				"user_dir,user_domain,user_domainroot,user_dir,oblog_user.userid,user_folder " & _
	
			"From oblog_log,oblog_user " & _
	 
			"Where oblog_user.userid = oblog_log.userid and ishide<>1 and isdraft=0 and passcheck=1 " & _

			"Order by iis Desc "
	
    Set rs = oblog.execute(msql)
   
    Do While Not rs.EOF
        
        Dim topic
        
		If rs("topic") <> "" Then
            topic = Replace(rs("topic"), "'", "")
            If topic <> "" Then
                If oblog.strLength(topic) > Int(l) Then
                    topic = oblog.InterceptStr(topic, Int(l) - 3) & "..."
                End If
            End If
        End If
		
        If oblog.setup(12, 0) = 1 Then
            userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
        Else
            userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
        End If
		
		
		'��ʾ������
		if (i mod 2 = 1) then
		
			show_toplog=show_toplog & _
              "<TR>"& _
                "<TD  width='169' height='26' >��"& _
				"<a href='" & userurl & "/" & oblog.filt_html(rs("logfile")) & "' title='" & oblog.filt_html(rs("topic")) & "'  target='_blank'>" & _
					oblog.filt_html(topic) & "</a>" & "(" & rs("iis") & ")"& _
				"</TD>"& _
              "</TR>"




'
'					 "<tr> " & _ 
'						"<td width='169' height='26' bgcolor='F3F3F3' style='padding-left:5px'>" & _
'						"  <img src='pic/left-1-2.gif' width='15' height='13' align='absmiddle'>" & _
'						"  <a href='" & userurl & "/" & oblog.filt_html(rs("logfile")) & "' title='" & oblog.filt_html(rs("topic")) & "'  target='_blank'>" & _
'						oblog.filt_html(topic) & "</a>" & "(" & rs("iis") & ")</td> "
		
			'��ӡ��һ��ʱ���ӡһ��
			'If i = 1 then show_toplog=show_toplog & "<td width='3' rowspan='10'></td> "
		
			'show_toplog=show_toplog & "</tr>"
		
		'��ʾż���У�����ʽ�������в�ͬ
		Else 
			show_toplog=show_toplog & _
              "<TR>"& _
                "<TD width='169' height='26' >��"& _
						  "<a href='" & userurl & "/" & oblog.filt_html(rs("logfile")) & "' title='" & oblog.filt_html(rs("topic")) & "'  target='_blank'>" & _
							oblog.filt_html(topic) & "</a>" & "(" & rs("iis") & ")" & _
				"</TD>"& _
              "</TR>"

			
			
			
			
			
'					 "<tr>" & _ 
'						"<td height='26' style='padding-left:5px'> " & _
'						"<img src='pic/left-1-3.gif' width='15' height='10' align='absmiddle'> " & _
'						  "<a href='" & userurl & "/" & oblog.filt_html(rs("logfile")) & "' title='" & oblog.filt_html(rs("topic")) & "'  target='_blank'>" & _
'							oblog.filt_html(topic) & "</a>" & "(" & rs("iis") & ")</td> " & _
'					 "</tr>"
		
		End If
		
        rs.MoveNext
		
		i = i + 1
 
    Loop
	
	rs.Close
    Set rs = Nothing
	
End Function


'*** End

Function show_userlog(userid,n, l, order,  subjectid,  info)
    Dim rs, strSql, strOrderSql, i, posttime,topic,strContent
    i = 0
    Select Case order
	    Case 1
	        'strOrderSql = " order by logid desc"
	        strOrderSql = " order by addtime desc" 
	    Case 2
	        strOrderSql = " order by iis desc"
	    Case 3
	        strOrderSql = " order by commentnum desc"
    End Select           
    
    strSql = "Select Top "&n&" topic,logfile,addtime,commentnum,iis,logid,author,userid"    
    strSql = strSql & " from oblog_log where ishide<>1 and isdraft=0 and passcheck=1 "
    strSql = strSql & "and blog_password=0  And (ispassword='' Or ispassword is null)"
    '���˵�����֮�������
    If is_Sqldata=1 Then
    	strSql = strSql & " And addtime<getdate() "
    Else
    	strSql = strSql & " And addtime<Now "
	End If
 	
 	If subjectid<>0 And IsNumeric(subjectid) Then
 		strSql = strSql & " And Subjectid=" & CLNG(subjectid)
 	End If
	strSql = strSql & " And UserId=" & CLNG(Userid) & strOrderSql
    'Response.Write strSql
    Set rs = oblog.execute(strSql)

    Do While Not rs.EOF
        strContent=strContent&"<li>"
        posttime = rs("addtime")
        If rs("topic") <> "" Then
            topic = Replace(rs("topic"), "'", "")
            If topic <> "" Then
                If oblog.strLength(topic) > Int(l) Then
                    topic = oblog.InterceptStr(topic, Int(l) - 3) & "..."
                End If
            End If
        End If
        strContent=strContent&"<a href=""go.asp?logid=" & rs("logid") &""" title="""&oblog.filt_html(rs(0))&""" target=_blank>"&oblog.filt_html(topic)&"</a>"
        If  CInt(info)=1 Then
            strContent=strContent&"("&formatdatetime(posttime,1)&")"
        End If
        strContent=strContent&"</li>"&vbcrlf
        rs.MoveNext
        i = i + 1
        If i >= Int(n) Then Exit Do
    Loop
    show_userlog=strContent
    strContent=""
    Set rs = Nothing
End Function

'���޸���
Function show_class(m)
   Dim rs
    '<tr> 
    '	<td>��<a href="#" target="_blank" class="a_k">���ͱ���</a></td>
	'	<td>��<a href="#" target="_blank" class="a_k">������ʳ</a></td>
    '</tr>
    Dim i, brstr
    show_class = ""
    m = Int(m)
    Set rs = oblog.execute("select id,classname from oblog_logclass Where idtype=0 order by RootID,OrderID")
    If m = 0 Then
        While Not rs.EOF
            show_class=show_class&"<a href=list.asp?classid="&rs(0)&">"&rs(1)&"</a><br/>"
            rs.MoveNext
        Wend
    Else
        i = 0
        Do While Not rs.EOF
		
			
			show_class=show_class&"<tr>" 
			
			show_class=show_class&"<td>��<a href=list.asp?classid="&rs(0)&">"&rs(1)&"</a></td>"
			
			'�ƶ�����һ����¼
			rs.MoveNext
			i = i + 1
			
        	If rs.EOF Then 
				show_class=show_class&"<td> </td>" & "<tr>"
				Exit Do
			Else 
				
				show_class=show_class&"<td>��<a href=list.asp?classid="&rs(0)&">"&rs(1)&"</a></td>" & "<tr>"
			End If
	
            'show_class=show_class&"<a href=list.asp?classid="&rs(0)&">"&rs(1)&"</a> "&brstr
            rs.MoveNext
			i = i + 1
        	If rs.EOF Then Exit Do
    	Loop
    End If
    Set rs = Nothing
End Function

Function show_comment(n, l)
    Dim rs
    set rs=oblog.execute("select top "&n&" mainid,commenttopic,comment_user,addtime,commentid from [oblog_comment] order by commentid desc")
    'show_comment="<ul>"
    While Not rs.EOF
        show_comment=show_comment&"<li><a href='go.asp?logid="&rs(0)&"&commentid="&rs(4)&"' target=_blank title='"&oblog.filt_html(rs(2))&"�ظ���"&rs(3)&"'>"&oblog.InterceptStr(oblog.filt_html(rs(1)),clng(l))&"</a></li>"
        rs.MoveNext
    Wend
    'show_comment=show_comment&"</ul>"
    Set rs = Nothing
End Function

'���޸���
Function show_subject(n)
    Dim i, rs
    i = 0
    'set rs=oblog.execute("select top "&n&" subjectid,oblog_subject.userid,subjectname,subjectlognum,user_dir,user_folder from [oblog_subject],oblog_user where oblog_subject.userid=oblog_user.userid and oblog_subject.oblog_subjecttype=0 order by subjectlognum desc")
    set rs=oblog.execute("Select a.*,b.username From (Select top " & n &" Subjectid,SubjectName,SubjectlogNum,userid From oBlog_subject where subjecttype=0 order by subjectlognum desc) a ,oblog_user b Where a.userid=b.userid and user_public<>1")
    'show_subject="<ul>"
    Do While Not rs.EOF
        
		
		'show_subject=show_subject&"<img src='pic/left-1-3.gif' width='15' height='10' vspace='7' align='absmiddle'>" & _
		show_subject=show_subject&"��" & _
								"<a href='blog.asp?name="&rs("username")&"&subjectid="&rs("subjectid")&"' target=_blank> "&oblog.filt_html(rs("subjectname"))&"("&rs("SubjectlogNum")&")</a>"
        '��Ϊ���һ����¼����
		If i < Int(n) Then  show_subject=show_subject&" <br>"
		
		rs.MoveNext
        i = i + 1
        If i >= Int(n) Then Exit Do
    Loop
    'show_subject=show_subject&"</ul>"
    Set rs = Nothing
End Function

'���޸���
Function show_blogupdate(n)
    Dim i, rs, userurl
    i = 0
    set rs=oblog.execute("select top "&n&" username,log_count,nickname,blogname,userid,user_dir,user_domain,user_domainroot,user_folder from [oblog_user] where lockuser=0 and user_public<>1 order by log_count desc")
    '��<a href="#" target="_blank" class="a_k">����ͳ�ƾֱ�����ͳ��</a><br>
    Do While Not rs.EOF
        If oblog.setup(12, 0) = 1 Then
            userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
        Else
            userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
        End If
        If rs("blogname") <> "" Then
            show_blogupdate=show_blogupdate&"�� <a href="&userurl&" target=_blank  class='a_k'>" & _
											rs("blogname")&"("&rs("log_count")&")</a>"
        Else
            show_blogupdate=show_blogupdate&"�� <a href="&userurl&" target=_blank class='a_k'>" & _
											rs("username")&"("&rs("log_count")&")</a>"
        End If
		
		'��Ϊ���һ����¼����
		If i < Int(n) Then  show_blogupdate=show_blogupdate&" <br>"
		
        rs.MoveNext
        i = i + 1
        If i >= Int(n) Then Exit Do
    Loop
    'show_blogupdate=show_blogupdate&"</ul>"
    Set rs = Nothing
End Function

'���޸��ˣ���ʾ�����û�
Function show_newblogger(n)
     Dim rs, userurl, i
	i = 0
    set rs=oblog.execute("select top "&n&" username,log_count,nickname,blogname,userid,user_dir,user_domain,user_domainroot,user_folder from [oblog_user] where lockuser=0 and user_public<>1 order by userid desc")
    'show_newblogger="<ul>"
    While Not rs.EOF
        If oblog.setup(12, 0) = 1 Then
            userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
        Else
            userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
        End If
        If rs(3) <> "" Then
            show_newblogger=show_newblogger&"��" & _
										"<a href="""&userurl&""" target=""_blank""> "&rs(3)&"("&rs(1)&")</a>"
            'show_newblogger=show_newblogger&"<img src='pic/left-1-3.gif' width='15' height='10' vspace='7' align='absmiddle'>" & _
										'"<a href="""&userurl&""" target=""_blank""> "&rs(3)&"("&rs(1)&")</a>"
        Else
            show_newblogger=show_newblogger&"��" & _
										"<a href="""&userurl&""" target=""_blank""> "&rs(0)&"("&rs(1)&")</a>"
            'show_newblogger=show_newblogger&"<img src='pic/left-1-3.gif' width='15' height='10' vspace='7' align='absmiddle'>" & _
										'"<a href="""&userurl&""" target=""_blank""> "&rs(0)&"("&rs(1)&")</a>"
        End If
		
		'��Ϊ���һ����¼����
		If i < Int(n) Then  show_newblogger=show_newblogger&" <br>"
        rs.MoveNext
		i = i + 1
        'If i >= Int(n) Then Exit Do
    Wend
    'show_newblogger=show_newblogger&"</ul>"
    Set rs = Nothing
End Function

'���޸��˵�
Function show_bestblog(n)
    Dim i, rs, userurl
    i = 0
    set rs=oblog.execute("select top "&n&" username,log_count,nickname,blogname,userid,user_dir,user_domain,user_domainroot,user_folder from [oblog_user] where user_isbest=1 order by log_count desc")
   
    Do While Not rs.EOF
        If oblog.setup(12, 0) = 1 Then
            userurl = "http://" & Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
        Else
            userurl = rs("user_dir") & "/" & rs("user_folder") & "/index." & f_ext
        End If
        If rs(3) <> "" Then
            show_bestblog=show_bestblog&"��" & _
										"<a href="""&userurl&""" target=""_blank""> "&rs(3)&"("&rs(1)&")</a>"
'			show_bestblog=show_bestblog&"<img src='pic/left-1-3.gif' width='15' height='10' vspace='7' align='absmiddle'>" & _
'										"<a href="""&userurl&""" target=""_blank""> "&rs(3)&"("&rs(1)&")</a>"
        Else
            show_bestblog=show_bestblog&"��" & _
										"<a href="""&userurl&""" target=""_blank""> "&rs(0)&"("&rs(1)&")</a>"
'			show_bestblog=show_bestblog&"<img src='pic/left-1-3.gif' width='15' height='10' vspace='7' align='absmiddle'>" & _
'										"<a href="""&userurl&""" target=""_blank""> "&rs(0)&"("&rs(1)&")</a>"
        End If
		
		'��Ϊ���һ����¼����
		If i < Int(n) Then  show_bestblog=show_bestblog&" <br>"
		
        rs.MoveNext
        i = i + 1
        If i >= Int(n) Then Exit Do
    Loop
    
    Set rs = Nothing

End Function

'���޸ĵ�
Function show_count()
    Dim rs,lngToday,lngYesteday
    If is_sqldata Then
        Set rs = oblog.execute("select count(logid) from oblog_log where datediff(d,truetime,getdate())=0")
        lngToday=rs(0)
        Set rs = oblog.execute("select count(logid) from oblog_log where datediff(d,truetime,getdate())=1")
        lngYesteday=rs(0)
    Else
        Set rs = oblog.execute("select count(logid) from oblog_log where datediff('d',truetime,now())=0")
         lngToday=rs(0)
        Set rs = oblog.execute("select count(logid) from oblog_log where datediff('d',truetime,now())=1")
        lngYesteday=rs(0)
    End If
    Set rs = Nothing
	
	' �� �û��� 26358 <br>
	
    show_count = "�� �û��� " & oblog.setup(24, 0) & "<br>"
    show_count=show_count&"�� ���£� "&oblog.setup(21,0)&"<br>"
    show_count=show_count&"�� ���ۣ� "&oblog.setup(22,0)&"<br>"
    show_count=show_count&"�� ���ԣ� "&oblog.setup(23,0)&"<br>"
    show_count=show_count&"�� ���գ� "&lngYesteday&"<br>"
    show_count=show_count&"�� ���գ� "&lngToday
   
End Function

Function show_sysxml()
    show_sysxml = "<a href='" & blogdir & "rss2.asp' target='_blank'><img src='Images/xml.gif' width='36' height='14' border='0'></a>"
End Function

Function show_friends()
    If IsNull(oblog.setup(19, 0)) Then
        show_friends = " "
    Else
        show_friends = oblog.setup(19, 0)
    End If
End Function

Function show_placard()
    If IsNull(oblog.setup(18, 0)) Then
        show_placard = " "
    Else
        show_placard = oblog.setup(18, 0)
    End If
End Function

Function show_bloger(m)
    Dim rs
    Dim i, brstr
    m = Int(m)
    Set rs = oblog.execute("select id,classname from oblog_userclass order by RootID,OrderID")
    If m = 0 Then
        While Not rs.EOF
            show_bloger=show_bloger&"<a href=listblogger.asp?usertype="&rs(0)&">"&rs(1)&"</a><br/>"
            rs.MoveNext
        Wend
    Else
        i = 0
        While Not rs.EOF
            i = i + 1
            If i = Int(m) Then
                brstr = "<br/>"
                i = 0
            Else
                brstr = ""
            End If
            show_bloger=show_bloger&"<a href=listblogger.asp?usertype="&rs(0)&">"&rs(1)&"</a> "&brstr
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
End Function

Function show_search(i)
    If i = 0 Then i = "" Else i = "<br />"
    show_search = "<form name='search' method='post' action='list.asp'>"
    show_search=show_search&"<select name='selecttype' id='selecttype'>"
    show_search=show_search&"<option value='topic' selected>���±���</option>"
    show_search=show_search&"<option value='logtext'>��������</option>"
    show_search=show_search&"<option value='id'>�û�����</option></select>"&i
    show_search=show_search&"<input name='keyword' type='text' id='keyword' size='16' maxlength='40'>"
    show_search=show_search&" <input type='submit' name='Submit' value='����'></form>"
End Function

'���޸���
Function show_cityblogger91(i)
    show_cityblogger91 = "<form name=""oblogform"" action=""listblogger.asp"">" & oblog.type_city91("", "") & _
						"<tr>" & _
						"<td height='24' align='center' valign='bottom' >" & _ 
						"<input name='Submit' type='image' src='pic/j-3-2.gif'  value='����' onclick='submit()' />" & _
						"</td></tr></form>"
						'<img src='pic/j-3-2.gif' width='44' height='17'></td></form>"
						'<input name="Submit" type="image" src="OblogStyle/OblogStyleAdminImages/login.gif" value="�ύ" onclick="submit()" />
						' <input type='submit' value='����'></form>"
    If i = 1 Then show_cityblogger91 = Replace(show_cityblogger91, "<select name='city'", "<br /><select name='city'")
End Function

 'ԭ�汾
Function show_cityblogger(i)
    show_cityblogger = "<form name=""oblogform"" action=""listblogger.asp"">" & oblog.type_city("", "") & " <input type='submit' value='����'></form>"
    If i = 1 Then show_cityblogger = Replace(show_cityblogger, "<select name='city'", "<br /><select name='city'")
End Function

'Function show_newphoto(n, i, w, h)
'    Dim rs, sReadMe,surl,imgsrc,fso
'	Set fso = server.CreateObject("Scripting.FileSystemObject")
'	
'    If i = 1 Then i = "<br />" Else i = ""
'    Set rs = oblog.execute("select top " & CLng(n) & " file_path,file_readme,oblog_upfile.userid,user_dir,username,nickname,logid from [oblog_user],oblog_upfile where oblog_user.userid=oblog_upfile.userid and isphoto=1 and ispower=0 order by fileid desc")
''��һ�������ͼƬ   wl   ... ... ��
'show_newphoto=show_newphoto&"<Tr align='center'>"
'
'	While Not rs.EOF
'	show_newphoto=show_newphoto&"<td>"
'        If IsNull(rs(1)) Then
'            sReadMe = ""
'        Else
'            sReadMe = oblog.filt_html(rs(1))
'        End If
'		if rs("logid")=0 or isnull(rs("logid")) then 
'			surl="<a href='"&rs("file_path")&"' target='_blank'>"
'		else
'			surl="<a href='more.asp?id="&rs("logid")&"' target='_blank'>"
'		end if
'		imgsrc=rs(0)
'		imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
'		imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
'		if  not fso.FileExists(Server.MapPath(imgsrc)) then
'			imgsrc=rs(0)
'		end if
'        show_newphoto=show_newphoto&"<a href='more.asp?id="&rs("logid")&"' target='_blank'><img src="""&imgsrc&""" width="""&w&""" height="""&h&""" hspace=""6"" border=""0"" vspace=""6"" alt='"& sReadMe &"' /></a>"&sReadMe &i
'        show_newphoto=show_newphoto&"</td>"
'		rs.MoveNext
'    Wend
'
'show_newphoto=show_newphoto&"</Tr>"
''��һ��������   wl   ... ... ��
''�ڶ��������ͼƬ�µ�˵������   wl   ... ... ��
'	rs.MoveFirst
'show_newphoto=show_newphoto&"<Tr align='center'>"
'
'	While Not rs.EOF
'	show_newphoto=show_newphoto&"<td>"
'        If IsNull(rs(1)) Then
'            sReadMe = ""
'        Else
'            sReadMe = oblog.filt_html(rs(1))
'        End If
'		if rs("logid")=0 or isnull(rs("logid")) then 
'			surl="<a href='"&rs("file_path")&"' target='_blank'>"
'		else
'			surl="<a href='more.asp?id="&rs("logid")&"' target='_blank'>"
'		end if
'		imgsrc=rs(0)
'		imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
'		imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
'		if  not fso.FileExists(Server.MapPath(imgsrc)) then
'			imgsrc=rs(0)
'		end if
'        show_newphoto=show_newphoto&"<a href='more.asp?id="&rs("logid")&"' target='_blank'><img src='pic/j-3-4.gif' hspace=""6"" border=""0"" vspace=""6"" alt='"& sReadMe &"' /></a>"&i
'		show_newphoto=show_newphoto&"</td>"
'        rs.MoveNext
'    Wend
'
'show_newphoto=show_newphoto&"</Tr>"
''�ڶ���������   wl   ... ... ��
'    Set rs = Nothing
'
'End Function
Function show_newphoto(n, i, w, h)
    Dim rs, sReadMe,surl,imgsrc,fso
	Set fso = server.CreateObject("Scripting.FileSystemObject")
    If i = 1 Then i = "<br />" Else i = ""
    Set rs = oblog.execute("select top " & CLng(n) & " file_path,file_readme,oblog_upfile.userid,user_dir,username,nickname,logid from [oblog_user],oblog_upfile where oblog_user.userid=oblog_upfile.userid and isphoto=1 and ispower=0 order by fileid desc")
    While Not rs.EOF
        If IsNull(rs(1)) Then
            sReadMe = ""
        Else
            sReadMe = oblog.filt_html(rs(1))
        End If
		if rs("logid")=0 or isnull(rs("logid")) then 
			surl="<a href='"&rs("file_path")&"' target='_blank'>"
		else
			surl="<a href='more.asp?id="&rs("logid")&"' target='_blank'>"
		end if
		imgsrc=rs(0)
		imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
		imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
		if  not fso.FileExists(Server.MapPath(imgsrc)) then
			imgsrc=rs(0)
		end if
        show_newphoto=show_newphoto&"<a href='more.asp?id="&rs("logid")&"' target='_blank'><img src="""&imgsrc&""" width="""&w&""" height="""&h&""" hspace=""6"" border=""0"" vspace=""6"" alt='"& sReadMe &"' /></a>"&i
        rs.MoveNext
    Wend
    Set rs = Nothing
End Function


Function show_blogstar()

    Dim rs
    Set rs = oblog.execute("select top 1 * from oblog_blogstar where ispass=1 order by id desc")
    If Not rs.EOF Then
		
        show_blogstar = "<div><a href='" & rs("userurl") & "' target='_blank'><img src=""" & rs("picurl") & """  hspace=""3"" border=""0"" vspace=""3"" alt='" & oblog.filt_html(rs("blogname")) & "' /></a></div>"
        show_blogstar=show_blogstar&"<div>�û���"&"<a href='"&rs("userurl")&"' target='_blank'>"&oblog.filt_html(rs("blogname"))&"</a></div>"
        show_blogstar=show_blogstar&"<div>��飺"&oblog.filt_html(rs("info"))&"</div>"
    Else
        show_blogstar = " "
    End If
    Set rs = Nothing
End Function

'���޸��������91health����ҳ
Public Function show_blogstar2(iNumber, iPerline, iWidth, iHeight)
    Dim rs, iCount, sLine
    If Not IsNumeric(iNumber) Then
        iNumber = 1
    Else
        iNumber = CLng(iNumber)
    End If
    'iWidth=160
    'iHeight=160
    If iNumber = 0 Then iNumber = 1
    Set rs = oblog.execute("select top " & iNumber & " * from oblog_blogstar where ispass=1 order by id desc")    
	
	If Not rs.EOF Then
		show_blogstar2 = "<tr align='center'>"
		 
		Do While Not rs.EOF
			sLine =  sLine & "<td>"  & "<a href='" & rs("userurl") & "' target='_blank'>" & _
							 "<img src='" & rs("picurl") & "' width='" & iWidth & "' height='" & iHeight & _
							 "' alt='" & oblog.filt_html(rs("info")) & "' class='border' border=0></a></td>"
						     '"<img src='" & rs("picurl") & "' width='94' height='98' class='border' border=0></a></td>"
		 	rs.MoveNext
		Loop
		
		show_blogstar2 = show_blogstar2 & sLine & "</tr>"
		
		'���������λ
		rs.MoveFirst
		
		show_blogstar2 =  show_blogstar2 &  "<tr align='center'>"
		
		sLine = ""
		Do While Not rs.EOF
			sLine =  sLine & "<td height='22'><img src='pic/j-3-4.gif' width='12' height='11' align='absmiddle' border=0>" & _ 
                  			"<a href='" & rs("userurl") & "' target='_blank'>" & oblog.filt_html(rs("blogname")) & "</a> </td>" 
			rs.MoveNext	
		Loop
		
		show_blogstar2 = show_blogstar2 & sLine & "</tr>"
		
	Else
        show_blogstar2 = " "
    End If
	
    rs.Close
    Set rs = Nothing
End Function


%>
