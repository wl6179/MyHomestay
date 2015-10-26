<%
Class class_sys
    Public cache_name, cache_name_custom, cache_data, Reloadtime, setup, userip, errstr, userdir, user_copyright, ver, is_password_cookies, is_gb2312
    Public logined_uname, logined_upass, logined_ulevel, logined_ulevel_team,logined_uisStop,logined_uisValidate, logined_ushowlogword, logined_uid, logined_uupfilemax, logined_uupfilesize, logined_udir, logined_isubb, logined_udomain, logined_ufolder,logined_uframe,logined_ucustomdomain
    Public comeurl, autoupdate
    Private Sub class_initialize()
        Reloadtime = 14400
        cache_name = blogdir& cache_name_user
        userip = request.ServerVariables("HTTP_X_FORWARDED_FOR")
        If userip = "" Then userip = request.ServerVariables("REMOTE_ADDR")
        comeurl = lcase(Trim(request.ServerVariables("HTTP_REFERER")))
        ver = "3.12"
        autoupdate = True '更新整站首页开关
        is_password_cookies = 1 '是否编码cookies,1为开启,0为关闭
        is_gb2312 = 1   '系统平台，1为简体中文平台，0为其他平台
    End Sub
    Public Property Let name(ByVal vNewValue)
        cache_name_custom = LCase(vNewValue)
    End Property
    Public Property Let Value(ByVal vNewValue)
        If cache_name_custom <> "" Then
            ReDim cache_data(2)
            cache_data(0) = vNewValue
            cache_data(1) = ServerDate(Now())'计算好时差；
            Application.Lock
            Application(cache_name & "_" & cache_name_custom) = cache_data
            Application.unLock
        Else
            Err.Raise vbObjectError + 1, "CacheServer", " please change the CacheName."
        End If
    End Property
    Public Property Get Value()
        If cache_name_custom <> "" Then
            cache_data = Application(cache_name & "_" & cache_name_custom)
            If IsArray(cache_data) Then
                Value = cache_data(0)
            Else
                Err.Raise vbObjectError + 1, "CacheServer", " The Cache_Data(" & cache_name_custom & ") Is Empty."
            End If
        Else
            Err.Raise vbObjectError + 1, "CacheServer", " please change the CacheName."
        End If
    End Property
    Public Function ObjIsEmpty()
        ObjIsEmpty = True
        cache_data = Application(cache_name & "_" & cache_name_custom)
        If Not IsArray(cache_data) Then Exit Function
        If Not IsDate(cache_data(1)) Then Exit Function
        If DateDiff("s", CDate(cache_data(1)), ServerDate(Now())) < (60 * Reloadtime) Then ObjIsEmpty = False
    End Function
    Public Sub DelCahe(MyCaheName)
        Application.Lock
        Application.Contents.Remove (cache_name & "_" & MyCaheName)
        Application.unLock
    End Sub
    Public Sub ReloadSetup()'是setup(0,*)提值的模型基础；
        Dim sql, rs, i
        sql = "select * from [oblog_setup] "
        Set rs = execute(sql)
        name = "setup"'是setup(0,*)提值的模型基础；
        Value = rs.GetRows(1)'是setup(0,*)提值的模型基础；
        Set rs = Nothing
        Application.Lock
        Application(cache_name & "index_update") = True
        Application(cache_name & "list_update") = True
        Application.unLock
    End Sub
    '读取用户目录及绑定的路径到缓存
    Public Sub ReloadUserdir()
        Dim sql, rs, s
        sql = "select userdir,dirdomain from [oblog_userdir] "
        Set rs = execute(sql)
        While Not rs.EOF
            s = s & rs(0) & "!!??((" & rs(1) & "##))=="
            rs.movenext
        Wend
        Application.Lock
        Application(cache_name & "dirdomain") = s
        Application.unLock
        Set rs = Nothing
    End Sub
    Public Sub start()
        name = "setup"
        If ObjIsEmpty() Then ReloadSetup()'是setup(0,*)提值的模型基础；
        setup = Value'通过触发ReloadSetup()函数，使得setup(0,*)模型基础已建立；
        user_copyright = setup(7, 0) & "<div id=""powered"" style=""text-align:center;""><a href=""http://www.myhomestay.com.cn"" target=""_blank""><img src=""images\meiqee.com.gif"" border=""0"" alt=""Powered by myhomestay.com.cn"" /></a></div>"
        If DateDiff("s", Application(oblog.cache_name & "index_updatetime"), ServerDate(Now())) > setup(17, 0) And Application(cache_name & "class_update") = False And autoupdate Then
            ReloadSetup()
            'ReloadUserdir()
            Application.Lock
            Application(cache_name & "index_update") = True
            Application(cache_name & "list_update") = True
            Application(cache_name & "class_update") = True
            Application.unLock
            response.Write ("<script src="""&blogdir&"index.asp?re=0""></script>")
        End If
    End Sub
    Public Sub sys_err(errmsg)
        Dim strErr
        strErr = strErr & "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
        strErr = strErr & "<link href='images/admin/admin_style.css' rel='stylesheet' type='text/css'></head><body>" & vbCrLf
        strErr = strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
        strErr = strErr & "  <br><tr align='center'><td height='22' class='title'><strong>错误信息</strong></td></tr>" & vbCrLf
        strErr = strErr & "  <tr><td height='100' class='tdbg' valign='top'><b>产生错误的可能原因：</b><br>" & errmsg & "</td></tr>" & vbCrLf
        strErr = strErr & "  <tr align='center'><td class='tdbg'><a href='javascript:history.go(-1)'>&lt;&lt; 返回上一页</a></td></tr>" & vbCrLf
        strErr = strErr & "</table>" & vbCrLf
        strErr = strErr & "</body></html>" & vbCrLf
        response.Write strErr
    End Sub
    Public Sub chk_comeurl()'检测是否有来源页面，或者是直接输入地址；
        If is_chk_comeurl = 1 Then
            Dim comeurl, curl
            comeurl = LCase(Trim(request.ServerVariables("HTTP_REFERER")))'请求的字符串內容(跳转的来 源页面 是什么)；
            If comeurl = "" Then
                response.Write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。</font></p>"
                response.End
            Else
                curl = Trim("http://" & request.ServerVariables("SERVER_NAME"))
                If Mid(comeurl, Len(curl) + 1, 1) = ":" Then
                    curl = curl & ":" & request.ServerVariables("SERVER_PORT")
                End If
                curl = LCase(curl & request.ServerVariables("SCRIPT_NAME"))'脚本的名字和包含路径；问题
                If LCase(Left(comeurl, InStrRev(comeurl, "/"))) <> LCase(Left(curl, InStrRev(curl, "/"))) Then
                    response.Write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许从外部链接地址访问本系统的后台管理页面。</font></p>"
                    response.End
                End If
            End If
        End If
    End Sub
    Public Function site_bottom()
        site_bottom = setup(7, 0) & vbCrLf
        site_bottom = site_bottom & "<div style=""display:none;clear: both;text-align: center;width: 100%;padding: 1;""><a href=""http://www.myhomestay.com.cn"" target=""_blank"">Powered by www.myhomestay.com.cn</a></div>" & vbCrLf
        site_bottom = site_bottom & "<script src=""login.asp?action=showindexlogin""></script><script src=""login.asp?action=showindexlogin_team""></script>"&vbCrLf&"</body>" & vbCrLf
        site_bottom = site_bottom & "</html>" & vbCrLf
    End Function
    Public Function execute(sql)
        If Not IsObject(conn) Then link_database
		'if instr(lcase(sql),"oblog_admin") then
			'response.Write("无权限"):response.End()
		'end if
        If is_debug = 0 Then
            On Error Resume Next
            Set execute = conn.execute(sql)
            If Err Then
                Err.Clear
                Set conn = Nothing
                response.Write "查询数据的时候发现错误，请检查您的查询代码是否正确。"
                response.End
            End If
        Else
            'Response.Write sql & "<br>"
            Set execute = conn.execute(sql)
        End If
    End Function
    Public Function chk_regname(regname)
        Dim regbadstr, i
        regbadstr = Split(setup(51, 0), vbCrLf)
        chk_regname = False
        For i = 0 To UBound(regbadstr)
            If Trim(regbadstr(i)) <> "" Then
                If Trim(regname) = Trim(regbadstr(i)) Then
                    chk_regname = True
                End If
            End If
            If chk_regname = True Then Exit For
        Next
    End Function
    Public Function chk_badword(Str)
        Dim badstr, i, n, tmp
		tmp=setup(50, 0)
		if instr(tmp,"$$$") then
			tmp=split(tmp,"$$$")
			badstr=tmp(1)
		else
			chk_badword=0
			exit function
		end if
        badstr = Split(badstr, vbCrLf)
        chk_badword = 0
        For i = 0 To UBound(badstr)
            If Trim(badstr(i)) <> "" Then
                If InStr(Str, Trim(badstr(i))) > 0 Then
                    chk_badword = 1
					exit function					
                End If
            End If
        Next
    End Function
    Public Function filt_badword(Str)
        Dim badstr,i,tmp
		tmp=setup(50, 0)
		if instr(tmp,"$$$") then
			tmp=split(tmp,"$$$")
			badstr=tmp(0)
		else
			badstr=tmp
		end if
        badstr = Split(badstr, vbCrLf)
        For i = 0 To UBound(badstr)
            If Trim(badstr(i)) <> "" Then
                Str = Replace(Str, badstr(i), "***")
            End If
        Next
        filt_badword = Str
    End Function
    Public Function getcode()
        getcode = "<img style='margin:0px;padding:0px;width:60px;height:18px;' src=""" & blogurl & "inc/code.asp"" style=""cursor:hand"" onclick=""this.src='"&blogurl&"inc/code.asp';"" alt=""看不清楚请点击一下验证码,MyHomestay"" />"
    End Function
    '检查验证码是否正确
    Public Function codepass()
        Dim CodeStr
        CodeStr = Trim(request("CodeStr"))'填写验证码的字符串；
        If CStr(Session("GetCode")) = CStr(CodeStr) And CodeStr <> "" Then'验证码图片一出现就立刻赋予的一个Session值("GetCode")；
            codepass = True
            'Session("GetCode")=empty
        Else
            codepass = False
            'Session("GetCode")=empty
        End If
    End Function
    Public Function type_domainroot(Str)
        Dim domainroot, i
        domainroot = Trim(oblog.setup(4, 0))
        If InStr(domainroot, "|") > 0 Then
            domainroot = Split(domainroot, "|")
            For i = 0 To UBound(domainroot)
                If Trim(domainroot(i)) <> "" Then
                    If domainroot(i) = Str Then
                    type_domainroot = type_domainroot & "<option value='" & Trim(domainroot(i)) & "' selected>" & "." & domainroot(i) & "</option>"
                    Else
                    type_domainroot = type_domainroot & "<option value='" & Trim(domainroot(i)) & "'>" & "." & domainroot(i) & "</option>"
                    End If
                End If
            Next
        Else
            type_domainroot = "<option value='" & domainroot & "'>" & "." & domainroot & "</option>"
        End If
    End Function
    Public Function show_class(kind, CurrentID, kindType)'kind是用户分类还是文章分类；文章分类中kindType是相册分类还是log分类；
        Dim rsClass, sqlClass, strTemp, tmpDepth, i
        Dim arrShowLine(20)
        For i = 0 To UBound(arrShowLine)
            arrShowLine(i) = False
        Next
        If kind = "user" Then
            sqlClass = "select * From oblog_userclass order by RootID,OrderID"
        ElseIf kind = "log" Then
            sqlClass = "select * From oblog_logclass  Where idType=" & kindType & " order by RootID,OrderID"
        End If
        Set rsClass = execute(sqlClass)
        If rsClass.bof And rsClass.EOF Then
            show_class = "<option value='0'>请选择类别</option>"
        Else
            show_class = "<option value='0'>请选择类别</option>"
            Do While Not rsClass.EOF
                tmpDepth = rsClass("Depth")
                If rsClass("NextID") > 0 Then
                    arrShowLine(tmpDepth) = True
                Else
                    arrShowLine(tmpDepth) = False
                End If
                    strTemp = "<option value='" & rsClass("id") & "'"
                If CurrentID > 0 And rsClass("id") = CurrentID Then'如果是当前id，则为选中状态；
                     strTemp = strTemp & " selected"
                End If
                strTemp = strTemp & ">"
                
                If tmpDepth > 0 Then'如果为 非顶级分类，则以此方式输出；
                    For i = 1 To tmpDepth
                        strTemp = strTemp & "&nbsp;&nbsp;"
                        If i = tmpDepth Then
                            If rsClass("NextID") > 0 Then
                                strTemp = strTemp & "├&nbsp;"
                            Else
                                strTemp = strTemp & "└&nbsp;"
                            End If
                        Else
                            If arrShowLine(i) = True Then
                                strTemp = strTemp & "│"
                            Else
                                strTemp = strTemp & "&nbsp;"
                            End If
                        End If
                    Next
                End If
                strTemp = strTemp & rsClass("classname")
                strTemp = strTemp & "</option>"
                show_class = show_class & strTemp
                rsClass.movenext
            Loop
        End If
        rsClass.Close
        Set rsClass = Nothing
    End Function
	Public Function show_class2(kind, CurrentID, kindType)'改为classid；
        Dim rsClass, sqlClass, strTemp, tmpDepth, i
		'''Response.Write("123456")
        Dim arrShowLine(20)
        For i = 0 To UBound(arrShowLine)
            arrShowLine(i) = False
        Next
        If kind = "user" Then
            sqlClass = "select * From oblog_userclass order by RootID,OrderID"
        ElseIf kind = "log" Then
            sqlClass = "select * From oblog_logclass  Where idType=" & kindType & " order by RootID,OrderID"
        End If
        Set rsClass = execute(sqlClass)
        If rsClass.bof And rsClass.EOF Then
            show_class2 = "<option value='0'>请选择类别</option>"
        Else
            show_class2 = "<option value='0'>请选择类别</option>"
            Do While Not rsClass.EOF
                tmpDepth = rsClass("Depth")
                If rsClass("NextID") > 0 Then
                    arrShowLine(tmpDepth) = True
                Else
                    arrShowLine(tmpDepth) = False
                End If
                    strTemp = "<option value='" & rsClass("classid") & "'"
                If CurrentID > 0 And rsClass("classid") = CurrentID Then'如果是当前id，则为选中状态；
                     strTemp = strTemp & " selected"
                End If
                strTemp = strTemp & ">"
                
                If tmpDepth > 0 Then'如果为 非顶级分类，则以此方式输出；
                    For i = 1 To tmpDepth
                        strTemp = strTemp & "&nbsp;&nbsp;"
                        If i = tmpDepth Then
                            If rsClass("NextID") > 0 Then
                                strTemp = strTemp & "├&nbsp;"
                            Else
                                strTemp = strTemp & "└&nbsp;"
                            End If
                        Else
                            If arrShowLine(i) = True Then
                                strTemp = strTemp & "│"
                            Else
                                strTemp = strTemp & "&nbsp;"
                            End If
                        End If
                    Next
                End If
                strTemp = strTemp & rsClass("classname")
                strTemp = strTemp & "</option>"
                show_class2 = show_class2 & strTemp
                rsClass.movenext
            Loop
        End If
        rsClass.Close
        Set rsClass = Nothing
    End Function
	
	Public Function show_class_Nav(kind, CurrentID, kindType, logclassid_Nav)'用于限制栏目空间的分类选择！WL;'kind是用户分类还是文章分类；文章分类中kindType是相册分类还是log分类；
        Dim rsClass, sqlClass, strTemp, tmpDepth, i
        Dim arrShowLine(20)
        For i = 0 To UBound(arrShowLine)
            arrShowLine(i) = False
        Next
        If kind = "user" Then
            sqlClass = "select * From oblog_userclass order by RootID,OrderID"
        ElseIf kind = "log" Then
            sqlClass = "select * From oblog_logclass Where classid="&logclassid_Nav&" And idType=" & kindType & " order by RootID,OrderID"
        End If
        Set rsClass = execute(sqlClass)
        If rsClass.bof And rsClass.EOF Then
            show_class_Nav = "<option value='0'>请选择类别</option>"
        Else
            show_class_Nav = "<option value='0'>请选择类别</option>"
            Do While Not rsClass.EOF
                tmpDepth = rsClass("Depth")
                If rsClass("NextID") > 0 Then
                    arrShowLine(tmpDepth) = True
                Else
                    arrShowLine(tmpDepth) = False
                End If
                    strTemp = "<option value='" & rsClass("classid") & "'"
                If CurrentID > 0 And rsClass("classid") = CurrentID Then'如果是当前id，则为选中状态；
                     strTemp = strTemp & " selected"
                End If
                strTemp = strTemp & ">"
                
                If tmpDepth > 0 Then'如果为 非顶级分类，则以此方式输出；
                    For i = 1 To tmpDepth
                        strTemp = strTemp & "&nbsp;&nbsp;"
                        If i = tmpDepth Then
                            If rsClass("NextID") > 0 Then
                                strTemp = strTemp & "├&nbsp;"
                            Else
                                strTemp = strTemp & "└&nbsp;"
                            End If
                        Else
                            If arrShowLine(i) = True Then
                                strTemp = strTemp & "│"
                            Else
                                strTemp = strTemp & "&nbsp;"
                            End If
                        End If
                    Next
                End If
                strTemp = strTemp & rsClass("classname")
                strTemp = strTemp & "</option>"
                show_class_Nav = show_class_Nav & strTemp
                rsClass.movenext
            Loop
        End If
        rsClass.Close
        Set rsClass = Nothing
    End Function
    Public Sub adderrstr(message)'加入一个错误信息；
        If errstr = "" Then
            errstr = message
        Else
            errstr = errstr & "_" & message
        End If
    End Sub
    Public Sub showerr()'输出错误信息；
        If errstr <> "" Then response.redirect "err.asp?message=" & errstr
    End Sub
    Public Sub showusererr()
        If errstr <> "" Then response.redirect "user_prompt.asp?message=" & errstr
    End Sub
	Public Sub showusererrHome()
        If errstr <> "" Then response.redirect "HomestayBackctrl_PromptHome.asp?message=" & errstr
    End Sub
	Public Sub showusererrCooperate()
        If errstr <> "" Then response.redirect "HomestayBackctrl_PromptCooperate.asp?message=" & errstr
    End Sub
	
    Public Sub SaveCookie(username, password, CookieDate, userurl)
        if cookies_domain<>"" then
			Response.Cookies(cookies_name).Domain=cookies_domain
		end if
        response.Cookies(cookies_name)("username") = CodeCookie(username)
        response.Cookies(cookies_name)("password") = CodeCookie(password)
        If userurl = "" Or userurl = "." Then userurl = " "
        response.Cookies(cookies_name)("userurl") = CodeCookie(userurl)
        Select Case CookieDate
            Case 0
                'not save
            Case 1
                response.Cookies(cookies_name).Expires = Date + 1'以天为单位；
            Case 2
                response.Cookies(cookies_name).Expires = Date + 31
            Case 3
                response.Cookies(cookies_name).Expires = Date + 365
        End Select
    End Sub
    Public Sub ob_chklogin(username, password, CookieDate)'ob菠萝阁用户；
        Dim rs, sql, userurl
        If Not IsObject(conn) Then link_database
        Set rs = server.CreateObject("adodb.recordset")
        sql = "select * from [oblog_user] where username='" & username & "' and password ='" & password & "' And user_level_team=0"
        rs.Open sql, conn, 1, 3
        If rs.bof And rs.EOF Then
            rs.Close: Set rs = Nothing
            adderrstr (username&":此家庭帐号不存在或者帐号密码不正确，请重新输入！"): showerr
            Exit Sub
        Else
            If rs("lockuser") = 1 Then
                rs.Close: Set rs = Nothing
                adderrstr ("对不起！你的家庭ID:"&username&"已被锁定，不能登陆！"): showerr
                Exit Sub
            Else
                rs("LastLoginIP") = oblog.userip
                rs("LastLoginTime") = ServerDate(Now())
                rs("LoginTimes") = rs("LoginTimes") + 1
                If Trim(oblog.setup(4, 0)) <> "" And oblog.setup(12, 0) = 1 Then
                    '启用二级域名
                    userurl = Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
                Else
                    '未启用二级域名则从根目录开始访问，不包含域名
                    'userurl= trim(oblog.setup(3,0)) & trim(rs("user_dir")) & "/" & trim(rs("userid")) & "/index." & f_ext
                    userurl = blogdir & Trim(rs("user_dir")) & "/" & Trim(rs("user_folder")) & "/index." & f_ext
                End If
                rs.Update
                SaveCookie username, password, CookieDate, userurl
                rs.Close: Set rs = Nothing
            End If
        End If
    End Sub
	Public Sub ob_chklogin_team(username, password, CookieDate)'合作伙伴check；
        Dim rs, sql, userurl
        If Not IsObject(conn) Then link_database
        Set rs = server.CreateObject("adodb.recordset")
        sql = "select * from [oblog_user] where username='" & username & "' and password ='" & password & "' And user_level_team=1"
        rs.Open sql, conn, 1, 3
        If rs.bof And rs.EOF Then
            rs.Close: Set rs = Nothing
            adderrstr (username&":合作帐号或密码错误，请重新输入！"): showerr
            Exit Sub
        Else
            If rs("lockuser") = 1 Then
                rs.Close: Set rs = Nothing
                adderrstr ("对不起！你的合作ID:"&username&"已被锁定，不能登陆！"): showerr
                Exit Sub
            Else
                rs("LastLoginIP") = oblog.userip
                rs("LastLoginTime") = ServerDate(Now())
                rs("LoginTimes") = rs("LoginTimes") + 1
                If Trim(oblog.setup(4, 0)) <> "" And oblog.setup(12, 0) = 1 Then
                    '启用二级域名
                    userurl = Trim(rs("user_domain")) & "." & Trim(rs("user_domainroot"))
                Else
                    '未启用二级域名则从根目录开始访问，不包含域名
                    'userurl= trim(oblog.setup(3,0)) & trim(rs("user_dir")) & "/" & trim(rs("userid")) & "/index." & f_ext
                    userurl = blogdir & Trim(rs("user_dir")) & "/" & Trim(rs("user_folder")) & "/index." & f_ext
                End If
                rs.Update
                SaveCookie username, password, CookieDate, userurl
                rs.Close: Set rs = Nothing
            End If
        End If
    End Sub
	
    Public Sub ot_chklogin(username, password, CookieDate)'ot指的是...从外部导入的用户表吧？ot_username是外部表的用户名e字段名，ot_password是外部表的密码e字段名；
        Dim sql, rs, rsreg
        If Not IsObject(ot_conn) Then link_database
        sql = "select * from " & ot_usertable & " where " & ot_username & "='" & username & "' and " & ot_password & " ='" & password & "'"
        Set rs = ot_conn.execute(sql)'ot_username是外部表的用户名e字段名，ot_password是外部表的密码e字段名；
        If rs.bof And rs.EOF Then
            Set rs = Nothing
            If IsObject(ot_conn) Then ot_conn.Close: Set ot_conn = Nothing
            oblog.adderrstr ("用户名或密码错误，请重新输入！"): oblog.showerr'加入一条错误信息，并且输出；
            Exit Sub
        Else'如果外部导入的表没有的情况下，那么就查询 菠萝阁本身的用户表！；
            Set rsreg = server.CreateObject("adodb.recordset")
            rsreg.Open "select * from [oblog_user] where username='" & username & "'", conn, 1, 3
            If rsreg.EOF Then
                Dim reguserlevel
                If oblog.setup(16, 0) = 1 Then reguserlevel = 6 Else reguserlevel = 7'如果注册不需要管理员审核（setup(16, 0) = 1），那么，reguserlevel = 7正式用户！；
                Set rsreg = server.CreateObject("adodb.recordset")
                rsreg.Open "select top 1 * from [oblog_user]", conn, 1, 3
                rsreg.addnew
                rsreg("username") = username
                rsreg("password") = "othertable"
                rsreg("user_dir") = oblog.setup(30, 0)
                rsreg("user_level") = reguserlevel
                rsreg("lockuser") = 0
                rsreg("en_blogteam") = 1
                rsreg("adddate") = ServerDate(Now())
                rsreg.Update
				oblog.execute("update oblog_user set user_folder=userid where username='"&username&"'")
                oblog.execute ("update oblog_setup set user_count=user_count+1")
                rsreg.Close
                Set rsreg = Nothing
                oblog.SaveCookie username, password, 0, " "
                oblog.CreateUserDir username, 1
                Set rs = Nothing
                oblog.showok "您是第一次激活blog系统，请完善blog资料!", "user_setting.asp"
                response.End()
            Else
                rsreg("LastLoginIP") = request.ServerVariables("REMOTE_ADDR")
                rsreg("LastLoginTime") = ServerDate(Now())
                rsreg("LoginTimes") = rsreg("LoginTimes") + 1
                rsreg.Update
            End If
            rsreg.Close
            Set rsreg = Nothing
            Set rs = Nothing
            If IsObject(ot_conn) Then ot_conn.Close: Set ot_conn = Nothing
            oblog.SaveCookie username, password, CookieDate, ""
        End If
    End Sub
    Public Function CheckUserLogined()
        Dim Logined, rsLogin, sqlLogin, ssql,user_info
        Logined = True
        logined_uname = filt_badstr(DecodeCookie(request.Cookies(cookies_name)("UserName")))
        logined_upass = filt_badstr(DecodeCookie(request.Cookies(cookies_name)("Password")))
		'response.Write(DecodeCookie(Request.Cookies(cookies_name)("UserName")))
		'response.End()
        If logined_uname = "" Then
            Logined = False
        End If
        If logined_upass = "" Then
            Logined = False
        End If
        ssql = "userid,user_level,user_showlogword_num,user_upfiles_max,user_upfiles_size,user_dir,isubbedit,user_domain,user_domainroot,lockuser,user_folder,user_info"&str_domain&",user_level_team,isStop,isValidate"
        If Logined = True Then
            If is_ot_user = 1 Then
                link_database
                sqlLogin = "select * from " & ot_usertable & " where " & ot_username & "='" & logined_uname & "' and " & ot_password & "='" & logined_upass & "'"
                Set rsLogin = ot_conn.execute(sqlLogin)
            Else
                sqlLogin = "select " & ssql & " from [oblog_user] where lockuser=0 and Username='" & logined_uname & "' and Password='" & logined_upass & "' and user_level_team=0"
                Set rsLogin = execute(sqlLogin)
            End If
            If rsLogin.bof And rsLogin.EOF Then'如果两者都没有这个帐号，登录失败；
                Logined = False
            Else'如果两者之中，有一个数据库中有此记录，那么... ...；
                If is_ot_user = 1 Then'在外接别用户数据库有这个用户的，那么...检测一下菠萝阁是否也有此帐户名了，...；
                    Set rsLogin = execute("select " & ssql & " from [oblog_user] where username='" & logined_uname & "' and user_level_team=0")
                End If
                
				If Not rsLogin.EOF Then'如果菠萝阁也有此帐户名，...就检索此帐号的各种参数；
                    If rsLogin(9) = 1 Then
                        Set rsLogin = Nothing
                        oblog.adderrstr ("当前用户已被系统锁定，无法进行操作，请联系管理员！")'先检测到此菠萝阁名是否已被锁定，那么就报错说此帐户不能使用！；
                        oblog.showerr
                    End If
                    logined_uid = rsLogin(0)
                    logined_ulevel = rsLogin(1)
					logined_ulevel_team = rsLogin("user_level_team")
					logined_uisStop = rsLogin("isStop")
					logined_uisValidate = rsLogin("isValidate")
					
                    logined_ushowlogword = rsLogin(2)
                    logined_uupfilemax = rsLogin(3)
                    logined_uupfilesize = rsLogin(4)
                    logined_udir = rsLogin(5)
                    logined_isubb = rsLogin(6)
                    logined_udomain = rsLogin(7) & "." & rsLogin(8)
                    logined_ufolder = rsLogin(10)
					if instr(rsLogin(11),"$") then'user_info字段；
						user_info=split(rsLogin(11),"$")
						logined_uframe=user_info(1)'$符号的后一个字符串1；
					else
						logined_uframe=1
					end if
					if true_domain=1 then 
						'判断用户绑定的顶级域名
						logined_ucustomdomain=rslogin("custom_domain")'绑定的顶级域名；
						if logined_ucustomdomain<>"" then
							logined_udomain=logined_ucustomdomain
						end if
					end if
                Else'如果菠萝阁没有此帐户名，...就创建此帐户；
                    Dim reguserlevel
                    If oblog.setup(16, 0) = 1 Then reguserlevel = 6 Else reguserlevel = 7
                    Dim rsreg
                    Set rsreg = server.CreateObject("adodb.recordset")
                    rsreg.Open "select top 1 * from [oblog_user]", conn, 1, 3
                    rsreg.addnew
                    rsreg("username") = logined_uname
                    rsreg("password") = "othertable"
                    rsreg("user_dir") = oblog.setup(30, 0)
                    rsreg("user_level") = reguserlevel
                    rsreg("lockuser") = 0
                    rsreg("en_blogteam") = 1
                    rsreg("adddate") = ServerDate(Now())
                    rsreg.Update
					oblog.execute("update oblog_user set user_folder=userid where username='"&logined_uname&"' and user_level_team=0")
                    execute ("update oblog_setup set user_count=user_count+1")
                    rsreg.Close
                    Set rsreg = Nothing
                    SaveCookie logined_uname, logined_upass, 0, " "
                    logined_ulevel = reguserlevel
					logined_ulevel_team = 0
                    oblog.CreateUserDir logined_uname, 1'创建一个该用户的新文件夹！！！；
                    Set rsLogin = Nothing
                    oblog.showok "您是第一次激活后台系统，请完善后台资料!", "user_setting.asp"
                    response.End()
                End If
            End If
            Set rsLogin = Nothing
        End If
        CheckUserLogined = Logined
    End Function
	Public Function CheckUserLogined_team()'是否为合作伙伴WL；
        Dim Logined, rsLogin, sqlLogin, ssql,user_info
        Logined = True
        logined_uname = filt_badstr(DecodeCookie(request.Cookies(cookies_name)("UserName")))
        logined_upass = filt_badstr(DecodeCookie(request.Cookies(cookies_name)("Password")))
		'response.Write(DecodeCookie(Request.Cookies(cookies_name)("UserName")))
		'response.End()
        If logined_uname = "" Then
            Logined = False
        End If
        If logined_upass = "" Then
            Logined = False
        End If
        ssql = "userid,user_level,user_showlogword_num,user_upfiles_max,user_upfiles_size,user_dir,isubbedit,user_domain,user_domainroot,lockuser,user_folder,user_info"&str_domain&",user_level_team,isStop,isValidate"
        If Logined = True Then
            If is_ot_user = 1 Then
                link_database
                sqlLogin = "select * from " & ot_usertable & " where " & ot_username & "='" & logined_uname & "' and " & ot_password & "='" & logined_upass & "'"
                Set rsLogin = ot_conn.execute(sqlLogin)
            Else
                sqlLogin = "select " & ssql & " from [oblog_user] where lockuser=0 and Username='" & logined_uname & "' and Password='" & logined_upass & "' and user_level_team=1"
                Set rsLogin = execute(sqlLogin)
            End If
            If rsLogin.bof And rsLogin.EOF Then'如果两者都没有这个帐号，登录失败；
                Logined = False
            Else'如果两者之中，有一个数据库中有此记录，那么... ...；
                If is_ot_user = 1 Then'在外接别用户数据库有这个用户的，那么...检测一下菠萝阁是否也有此帐户名了，...；
                    Set rsLogin = execute("select " & ssql & " from [oblog_user] where username='" & logined_uname & "' and user_level_team=1")
                End If
                
				If Not rsLogin.EOF Then'如果菠萝阁也有此帐户名，...就检索此帐号的各种参数；
                    If rsLogin(9) = 1 Then
                        Set rsLogin = Nothing
                        oblog.adderrstr ("当前合作伙伴用户已被系统锁定，无法进行操作，请联系管理员！")'先检测到此菠萝阁名是否已被锁定，那么就报错说此帐户不能使用！；
                        oblog.showerr
                    End If
                    logined_uid = rsLogin(0)
                    logined_ulevel = rsLogin(1)
					logined_ulevel_team = rsLogin("user_level_team")'等于1；
					logined_uisStop = rsLogin("isStop")
					logined_uisValidate = rsLogin("isValidate")
					
                    logined_ushowlogword = rsLogin(2)
                    logined_uupfilemax = rsLogin(3)
                    logined_uupfilesize = rsLogin(4)
                    logined_udir = rsLogin(5)
                    logined_isubb = rsLogin(6)
                    logined_udomain = rsLogin(7) & "." & rsLogin(8)
                    logined_ufolder = rsLogin(10)
					if instr(rsLogin(11),"$") then'user_info字段；
						user_info=split(rsLogin(11),"$")
						logined_uframe=user_info(1)'$符号的后一个字符串1；
					else
						logined_uframe=1
					end if
					if true_domain=1 then 
						'判断用户绑定的顶级域名
						logined_ucustomdomain=rslogin("custom_domain")'绑定的顶级域名；
						if logined_ucustomdomain<>"" then
							logined_udomain=logined_ucustomdomain
						end if
					end if
                Else'如果菠萝阁没有此帐户名，...就创建此帐户；
                    Dim reguserlevel
                    If oblog.setup(16, 0) = 1 Then reguserlevel = 6 Else reguserlevel = 7
                    Dim rsreg
                    Set rsreg = server.CreateObject("adodb.recordset")
                    rsreg.Open "select top 1 * from [oblog_user]", conn, 1, 3
                    rsreg.addnew
                    rsreg("username") = logined_uname
                    rsreg("password") = "othertable"
                    rsreg("user_dir") = oblog.setup(30, 0)
                    rsreg("user_level") = reguserlevel
                    rsreg("lockuser") = 0
                    rsreg("en_blogteam") = 1
                    rsreg("adddate") = ServerDate(Now())
                    rsreg.Update
					oblog.execute("update oblog_user set user_folder=userid where username='"&logined_uname&"' and user_level_team=1")
                    execute ("update oblog_setup set user_count=user_count+1")
                    rsreg.Close
                    Set rsreg = Nothing
                    SaveCookie logined_uname, logined_upass, 0, " "

                    logined_ulevel = reguserlevel
					logined_ulevel_team = 0
                    oblog.CreateUserDir logined_uname, 1'创建一个该用户的新文件夹！！！；
                    Set rsLogin = Nothing
                    oblog.showok "您是第一次激活后台系统，请完善后台资料!", "user_setting.asp"
                    response.End()
                End If
            End If
            Set rsLogin = Nothing
        End If
        CheckUserLogined_team = Logined
		'CheckUserLogined = False
    End Function
	Public Function CheckUserLogined_home()'是否为家庭用户WL；
        Dim Logined, rsLogin, sqlLogin, ssql,user_info
        Logined = True
        logined_uname = filt_badstr(DecodeCookie(request.Cookies(cookies_name)("UserName")))
        logined_upass = filt_badstr(DecodeCookie(request.Cookies(cookies_name)("Password")))
		'response.Write(DecodeCookie(Request.Cookies(cookies_name)("UserName")))
		'response.End()
        If logined_uname = "" Then
            Logined = False
        End If
        If logined_upass = "" Then
            Logined = False
        End If
        ssql = "userid,user_level,user_showlogword_num,user_upfiles_max,user_upfiles_size,user_dir,isubbedit,user_domain,user_domainroot,lockuser,user_folder,user_info"&str_domain&",user_level_team,isStop,isValidate"
        If Logined = True Then
            If is_ot_user = 1 Then
                link_database
                sqlLogin = "select * from " & ot_usertable & " where " & ot_username & "='" & logined_uname & "' and " & ot_password & "='" & logined_upass & "'"
                Set rsLogin = ot_conn.execute(sqlLogin)
            Else
                sqlLogin = "select " & ssql & " from [oblog_user] where lockuser=0 and Username='" & logined_uname & "' and Password='" & logined_upass & "' and user_level_team=0 And user_level=6"
                Set rsLogin = execute(sqlLogin)
            End If
            If rsLogin.bof And rsLogin.EOF Then'如果两者都没有这个帐号，登录失败；
                Logined = False
            Else'如果两者之中，有一个数据库中有此记录，那么... ...；
                If is_ot_user = 1 Then'在外接别用户数据库有这个用户的，那么...检测一下菠萝阁是否也有此帐户名了，...；
                    Set rsLogin = execute("select " & ssql & " from [oblog_user] where username='" & logined_uname & "' and user_level_team=0 And user_level=6")
                End If
                
				If Not rsLogin.EOF Then'如果菠萝阁也有此帐户名，...就检索此帐号的各种参数；
                    If rsLogin(9) = 1 Then
                        Set rsLogin = Nothing
                        oblog.adderrstr ("当前家庭用户已被系统锁定，无法进行操作，请联系管理员！")'先检测到此菠萝阁名是否已被锁定，那么就报错说此帐户不能使用！；
                        oblog.showerr
                    End If
                    logined_uid = rsLogin(0)
                    logined_ulevel = rsLogin(1)
					logined_ulevel_team = rsLogin("user_level_team")'等于1；
					logined_uisStop = rsLogin("isStop")
					logined_uisValidate = rsLogin("isValidate")
					
                    logined_ushowlogword = rsLogin(2)
                    logined_uupfilemax = rsLogin(3)
                    logined_uupfilesize = rsLogin(4)
                    logined_udir = rsLogin(5)
                    logined_isubb = rsLogin(6)
                    logined_udomain = rsLogin(7) & "." & rsLogin(8)
                    logined_ufolder = rsLogin(10)
					if instr(rsLogin(11),"$") then'user_info字段；
						user_info=split(rsLogin(11),"$")
						logined_uframe=user_info(1)'$符号的后一个字符串1；
					else
						logined_uframe=1
					end if
					if true_domain=1 then 
						'判断用户绑定的顶级域名
						logined_ucustomdomain=rslogin("custom_domain")'绑定的顶级域名；
						if logined_ucustomdomain<>"" then
							logined_udomain=logined_ucustomdomain
						end if
					end if
                Else'如果菠萝阁没有此帐户名，...就创建此帐户；
                    Dim reguserlevel
                    If oblog.setup(16, 0) = 1 Then reguserlevel = 6 Else reguserlevel = 7
                    Dim rsreg
                    Set rsreg = server.CreateObject("adodb.recordset")
                    rsreg.Open "select top 1 * from [oblog_user]", conn, 1, 3
                    rsreg.addnew
                    rsreg("username") = logined_uname
                    rsreg("password") = "othertable"
                    rsreg("user_dir") = oblog.setup(30, 0)
                    rsreg("user_level") = reguserlevel
                    rsreg("lockuser") = 0
                    rsreg("en_blogteam") = 1
                    rsreg("adddate") = ServerDate(Now())
                    rsreg.Update
					oblog.execute("update oblog_user set user_folder=userid where username='"&logined_uname&"' and user_level_team=0 And user_level=6")
                    execute ("update oblog_setup set user_count=user_count+1")
                    rsreg.Close
                    Set rsreg = Nothing
                    SaveCookie logined_uname, logined_upass, 0, " "

                    logined_ulevel = reguserlevel
					logined_ulevel_team = 0
                    oblog.CreateUserDir logined_uname, 1'创建一个该用户的新文件夹！！！；
                    Set rsLogin = Nothing
                    oblog.showok "您是第一次激活家庭帐号，请完善后台资料!", "HomestayBackctrl-indexHome.asp"
                    response.End()
                End If
            End If
            Set rsLogin = Nothing
        End If
        CheckUserLogined_home = Logined
		'CheckUserLogined = False
    End Function
	
    Public Sub CreateUserDir(ustr, action)'===初始化一个菠萝阁的新文件夹和各个新的基本文件；===
        Dim fso, sql, rs, udir, uid, upath, loginstr, searchstr, bname, ufolder,utruepath
        sql = "select userid,user_dir,blogname,user_folder,user_domain,user_domainroot"&str_domain&" from oblog_user where "
        If action = 0 Then sql = sql & "userid=" & CLng(ustr) Else sql = sql & "username='" & filt_badstr(ustr) & "'"
        Set rs = execute(sql)
        If Not rs.EOF Then
            udir = rs(1)
            uid = rs(0)
            bname = rs(2)
            ufolder = rs(3)
			if true_domain=1 then'某个菠萝阁所在的文件夹；
				if rs("custom_domain")<>"" and not isnull(rs("custom_domain")) then
					utruepath="http://"&rs("custom_domain")&"/"
				else
					utruepath="http://"&rs("user_domain")&"."&rs("user_domainroot")&"/"
				end if
			else
				utruepath=blogdir & udir & "/" & ufolder & "/"
			end if
            If bname = "" Or IsNull(bname) Then bname = " "
            '''searchstr = "<form name='search' method='post' action='" & blogurl & "HomestayNav.asp?userid=" & uid & "' target=""_blank"">"
			searchstr = "<form name='search' method='post' action='" & blogurl & "HomestayNav.asp' target=""_blank"">"
            searchstr = searchstr & "<select name='selecttype' id='selecttype'>"
            searchstr = searchstr & "<option value='topic' selected>文章标题</option>"
            searchstr = searchstr & "<option value='logtext'>文章内容</option></select><br />"
            searchstr = searchstr & "<input name='keyword' type='text' id='keyword' size='16' maxlength='40'>"
            searchstr = searchstr & " <input type='submit' name='Submit' value='搜索'></form>"
            upath = server.MapPath(udir)'对根目录下的u文件夹的磁盘路径；
            Set fso = server.CreateObject("scripting.filesystemobject")
            If fso.FolderExists(upath) = False Then fso.CreateFolder (upath)
            upath = server.MapPath(blogdir & udir & "/" & ufolder)
            If fso.FolderExists(upath) = False Then fso.CreateFolder (upath)'如果不存在这么一个能存放一批菠萝阁的u文件夹，那么，就即时创建它！；
            Call oblog.BuildFile(upath & "/index." & f_ext, "暂无文章,请发布文章或者更新栏目首页!")
            Call oblog.BuildFile(upath & "/message." & f_ext, "暂无留言,请更新栏目留言板!")
            Call oblog.BuildFile(upath & "/photo." & f_ext, "暂无相册,请添加图片或更新栏目相册首页!")
            upath = server.MapPath(blogdir & udir & "/" & ufolder & "/calendar")
            If fso.FolderExists(upath) = False Then fso.CreateFolder (upath)
            If f_ext = "htm" Or f_ext = "html" Then
                Call oblog.BuildFile(upath & "/0.htm", oblog.htm2js_div(" ", "calendar"))
            Else
                Call oblog.BuildFile(upath & "/0.htm", " ")
            End If
            upath = server.MapPath(blogdir & udir & "/" & ufolder & "/inc")'为菠萝阁创建inc文件夹，并写入各种可调用的 标签！wl；
            If fso.FolderExists(upath) = False Then fso.CreateFolder (upath)
            If f_ext = "htm" Or f_ext = "html" Then
                Call oblog.BuildFile(upath & "/show_blogname.htm", oblog.htm2js_div(filt_html(bname), "blogname"))'生成js片段：之菠萝阁名；
                Call oblog.BuildFile(upath & "/show_placard.htm", oblog.htm2js_div(" ", "placard"))'生成js片段：之公告栏的html代码（通过js语句innerHTML，插入公告栏的html代码给<标签id="placard" >的页面标签！）；
                Call oblog.BuildFile(upath & "/show_subject.htm", oblog.htm2js_div(" ", "subject"))'生成js片段：之菠萝阁自己分类；
                Call oblog.BuildFile(upath & "/show_newblog.htm", oblog.htm2js_div(" ", "newblog"))'生成js片段：之最新log；
                Call oblog.BuildFile(upath & "/show_comment.htm", oblog.htm2js_div(" ", "comment"))'生成js片段：之最新回复；
                Call oblog.BuildFile(upath & "/show_links.htm", oblog.htm2js_div(" ", "links"))'生成js片段：之友情链接；
                Call oblog.BuildFile(upath & "/show_info.htm", oblog.htm2js_div(" ", "info"))'生成js片段：之统计信息；
                Call oblog.BuildFile(upath & "/show_search.htm", oblog.htm2js_div(searchstr, "search"))'生成js片段：之搜索栏；
                Call oblog.BuildFile(upath & "/show_newmessage.htm", oblog.htm2js_div("<a href=""" & utruepath & "message." & f_ext & "#cmt"">::签写留言::</a> ", "newmessage"))'生成js片段：之留言；
            Else
                Call oblog.BuildFile(upath & "/show_blogname.htm", filt_html(bname))
                Call oblog.BuildFile(upath & "/show_placard.htm", " ")
                Call oblog.BuildFile(upath & "/show_subject.htm", " ")
                Call oblog.BuildFile(upath & "/show_newblog.htm", " ")
                Call oblog.BuildFile(upath & "/show_comment.htm", " ")
                Call oblog.BuildFile(upath & "/show_links.htm", " ")
                Call oblog.BuildFile(upath & "/show_info.htm", " ")
                Call oblog.BuildFile(upath & "/show_search.htm", searchstr)
                Call oblog.BuildFile(upath & "/show_newmessage.htm", "<a href=""" & utruepath& "message." & f_ext & "#cmt"">::签写留言::</a> ")
            End If
			if logfilepath=1 then
				upath = server.MapPath(blogdir & udir & "/" & ufolder & "/archives")
				If fso.FolderExists(upath) = False Then fso.CreateFolder (upath)
			end if
            Set fso = Nothing
            Set rs = Nothing
        Else
            Set rs = Nothing
            response.Write ("没找到该用户，无法建立目录。")
            Exit Sub
        End If
    End Sub
    Public Sub showok(Str, url)
        url = Trim(url)
        If url <> "" Then
            response.Write "<script language=JavaScript>alert(""" & Str & """);window.location='" & url & "'</script>"
        Else
            If comeurl = "" Then
                response.Write "<script language=JavaScript>alert(""" & Str & """);history.go(-1)</script>"
            Else
                response.Write "<script language=JavaScript>alert(""" & Str & """);window.location='" & comeurl & "'</script>"
            End If
        End If
    End Sub
    Public Function type_city(province, city)
        Dim tmpstr
        tmpstr = "<select onchange=setcity(); name='province' >"
        tmpstr = tmpstr & "<option value=''>请选择省份</option>"
        tmpstr = tmpstr & "<option "
        tmpstr = tmpstr & "value=北京>北京</option> <option value=上海>上海</option> "
        tmpstr = tmpstr & "<option value=重庆>重庆</option> <option "
        tmpstr = tmpstr & "value=福建>福建</option> <option value=甘肃>甘肃</option> "
        tmpstr = tmpstr & "<option value=广东>广东</option> <option "
        tmpstr = tmpstr & "value=广西>广西</option> <option value=贵州>贵州</option> "
        tmpstr = tmpstr & "<option value=海南>海南</option> <option "
        tmpstr = tmpstr & "value=河北>河北</option> <option value=黑龙江>黑龙江</option> "
        tmpstr = tmpstr & "<option value=河南>河南</option> <option "
        tmpstr = tmpstr & "value=香港>香港</option> <option value=湖北>湖北</option> "
        tmpstr = tmpstr & "<option value=湖南>湖南</option> <option "
        tmpstr = tmpstr & "value=江苏>江苏</option> <option value=江西>江西</option> "
        tmpstr = tmpstr & "<option value=吉林>吉林</option> <option "
        tmpstr = tmpstr & "value=辽宁>辽宁</option> <option value=澳门>澳门</option>"
        tmpstr = tmpstr & "<option value=内蒙古>内蒙古</option> <option "
        tmpstr = tmpstr & "value=宁夏>宁夏</option> <option value=青海>青海</option> "
        tmpstr = tmpstr & "<option value=山东>山东</option> <option "
        tmpstr = tmpstr & "value=安徽>安徽</option> <option value=山西>山西</option> "
        tmpstr = tmpstr & "<option value=陕西>陕西</option> <option "
        tmpstr = tmpstr & "value=四川>四川</option> <option value=台湾>台湾</option> "
        tmpstr = tmpstr & "<option value=天津>天津</option> <option "
        tmpstr = tmpstr & "value=新疆>新疆</option> <option value=西藏>西藏</option> "
        tmpstr = tmpstr & "<option value=云南>云南</option> <option "
        tmpstr = tmpstr & "value=浙江>浙江</option> <option "
        tmpstr = tmpstr & "value=海外>海外</option></select>"
        tmpstr = tmpstr & " <select name='city' >"
        tmpstr = tmpstr & "</select>"
        tmpstr = tmpstr & "<SCRIPT src=""inc/getcity.js""></SCRIPT>"
        tmpstr = tmpstr & "<SCRIPT>initprovcity('" & province & "','" & city & "');</SCRIPT>"
        type_city = tmpstr
    End Function
	Public Function type_citySelectSubmit(province, city)'专为选择onchange submit作用的函数；
        Dim tmpstr
        tmpstr = "<select onchange=setcity();submit(); name='province' >"
        tmpstr = tmpstr & "<option value=''>请选择省份</option>"
        tmpstr = tmpstr & "<option "
        tmpstr = tmpstr & "value=北京>北京</option> <option value=上海>上海</option> "
        tmpstr = tmpstr & "<option value=重庆>重庆</option> <option "
        tmpstr = tmpstr & "value=福建>福建</option> <option value=甘肃>甘肃</option> "
        tmpstr = tmpstr & "<option value=广东>广东</option> <option "
        tmpstr = tmpstr & "value=广西>广西</option> <option value=贵州>贵州</option> "
        tmpstr = tmpstr & "<option value=海南>海南</option> <option "
        tmpstr = tmpstr & "value=河北>河北</option> <option value=黑龙江>黑龙江</option> "
        tmpstr = tmpstr & "<option value=河南>河南</option> <option "
        tmpstr = tmpstr & "value=香港>香港</option> <option value=湖北>湖北</option> "
        tmpstr = tmpstr & "<option value=湖南>湖南</option> <option "
        tmpstr = tmpstr & "value=江苏>江苏</option> <option value=江西>江西</option> "
        tmpstr = tmpstr & "<option value=吉林>吉林</option> <option "
        tmpstr = tmpstr & "value=辽宁>辽宁</option> <option value=澳门>澳门</option>"
        tmpstr = tmpstr & "<option value=内蒙古>内蒙古</option> <option "
        tmpstr = tmpstr & "value=宁夏>宁夏</option> <option value=青海>青海</option> "
        tmpstr = tmpstr & "<option value=山东>山东</option> <option "
        tmpstr = tmpstr & "value=安徽>安徽</option> <option value=山西>山西</option> "
        tmpstr = tmpstr & "<option value=陕西>陕西</option> <option "
        tmpstr = tmpstr & "value=四川>四川</option> <option value=台湾>台湾</option> "
        tmpstr = tmpstr & "<option value=天津>天津</option> <option "
        tmpstr = tmpstr & "value=新疆>新疆</option> <option value=西藏>西藏</option> "
        tmpstr = tmpstr & "<option value=云南>云南</option> <option "
        tmpstr = tmpstr & "value=浙江>浙江</option> <option "
        tmpstr = tmpstr & "value=海外>海外</option></select>"
        tmpstr = tmpstr & " <select name='city' onchange='submit();'>"
        tmpstr = tmpstr & "</select>"
        tmpstr = tmpstr & "<SCRIPT src=""inc/getcityArea.js""></SCRIPT>"
        tmpstr = tmpstr & "<SCRIPT>initprovcity('" & province & "','" & city & "');</SCRIPT>"
        type_citySelectSubmit = tmpstr
    End Function
	
    Public Sub type_job(job)
        Dim tmpstr
        tmpstr = "<select name='job' id='job'><option value=''>----请选择职业----<option value='财会/金融'> 财会/金融<option value='工程师'  > 工程师<option value='顾问'  > 顾问<option value='计算机相关行业(Internet)'  > 计算机相关行业 (Internet)<option value='计算机相关行业（其他）'  > 计算机相关行业（其他）<option value='家庭主妇'  > 家庭主妇<option value='教育/培训'  > 教育/培训<option value='客户服务/支持'  > 客户服务/支持<option value='零售商/手工工人'  > 零售商/手工工人<option value='退休'  > 退休<option value='无职业'  > 无职业<option value='销售/市场/广告'  > 销售/市场/广告<option value='学生'  > 学生<option value='研究和开发'  > 研究和开发<option value='一般管理'  > 一般管理<option value='政府/军队'  > 政府/军队<option value='执行官/高级管理'  > 执行官/高级管理<option value='制造/生产/操作'  > 制造/生产/操作<option value='专业人员（医药、法律等）'  > 专业人员（医药、法律等）<option value='自雇/业主'  > 自雇/业主<option value='其他'  > 其他</select>"
        response.Write (tmpstr)
%>
<SCRIPT language=javascript>
    var jobObject = document.oblogform["job"];
    for(var i = 0; i < document.oblogform["job"].options.length; i++) {
        if (document.oblogform["job"].options[i].value=="<%=trim(job)%>")
        {
            document.oblogform["job"].selectedIndex = i;
        }
    }
</SCRIPT>
<%
    End Sub
	Public Function type_job2(job)
        Dim tmpstr
        tmpstr = "<select name='job' id='job'><option value=''>----请选择职业----<option value='财会/金融'> 财会/金融<option value='工程师'  > 工程师<option value='顾问'  > 顾问<option value='计算机相关行业(Internet)'  > 计算机相关行业 (Internet)<option value='计算机相关行业（其他）'  > 计算机相关行业（其他）<option value='家庭主妇'  > 家庭主妇<option value='教育/培训'  > 教育/培训<option value='客户服务/支持'  > 客户服务/支持<option value='零售商/手工工人'  > 零售商/手工工人<option value='退休'  > 退休<option value='无职业'  > 无职业<option value='销售/市场/广告'  > 销售/市场/广告<option value='学生'  > 学生<option value='研究和开发'  > 研究和开发<option value='一般管理'  > 一般管理<option value='政府/军队'  > 政府/军队<option value='执行官/高级管理'  > 执行官/高级管理<option value='制造/生产/操作'  > 制造/生产/操作<option value='专业人员（医药、法律等）'  > 专业人员（医药、法律等）<option value='自雇/业主'  > 自雇/业主<option value='其他'  > 其他</select>"
        

tmpstr = tmpstr&"<SCRIPT language=javascript>"
    tmpstr = tmpstr&"var jobObject = document.oblogform['job'];"
    tmpstr = tmpstr&"for(var i = 0; i < document.oblogform['job'].options.length; i++) {"
        tmpstr = tmpstr&"if (document.oblogform['job'].options[i].value=='" &trim(job)&"')"
        tmpstr = tmpstr&"{"
            tmpstr = tmpstr&"document.oblogform['job'].selectedIndex = i;"
        tmpstr = tmpstr&"}"
    tmpstr = tmpstr&"}"
tmpstr = tmpstr&"</SCRIPT>"
	type_job2 = tmpstr
    End Function
    Public Function type_dateselect(addtime, n)
        Dim y, m, d, ttime
        If addtime = "" Then ttime = ServerDate(Now()) Else ttime = addtime
        type_dateselect = type_dateselect& ("<select name=selecty" & n & ">") & vbCrLf
        For y = 1900 To 2010
            If Year(ttime) = y Then
                type_dateselect = type_dateselect& "<option value=" & y & " selected>" & y & "年</option>" & vbCrLf
            Else
                type_dateselect = type_dateselect& "<option value=" & y & ">" & y & "年</option>" & vbCrLf
            End If
        Next
        type_dateselect = type_dateselect& "</select><select name=selectm" & n & ">" & vbCrLf
        For m = 1 To 12
            If Month(ttime) = m Then
                type_dateselect = type_dateselect& "<option value=" & m & " selected>" & m & "月</option>" & vbCrLf
            Else
                type_dateselect = type_dateselect& "<option value=" & m & ">" & m & "月</option>" & vbCrLf
            End If
        Next
        type_dateselect = type_dateselect& ("</select><select name=selectd" & n & ">") & vbCrLf
        For d = 1 To 31
            If Day(ttime) = d Then
                type_dateselect = type_dateselect& "<option value=" & d & " selected>" & d & "日</option>" & vbCrLf
            Else
                type_dateselect = type_dateselect& "<option value=" & d & ">" & d & "日</option>" & vbCrLf
            End If
        Next
        type_dateselect = type_dateselect& ("</select>") & vbCrLf
    End Function
    Public Sub chk_commenttime()
        Dim lasttime
        lasttime = Session("chk_commenttime" & Replace(request.ServerVariables("REMOTE_ADDR"), ".", ""))
        If IsDate(lasttime) Then
            If DateDiff("s", lasttime, ServerDate(Now())) < setup(49, 0) Then
                response.Write ("<script language=javascript>alert('" & setup(49, 0) & "秒后才能回复或评论。');window.history.back(-1);</script>")
                response.End
            End If
        End If
    End Sub
    Public Function filtpath(Str)
        If is_relativepath = 1 Then'如果是相对路径；
            Dim nurl
            nurl = Trim("http://" & request.ServerVariables("SERVER_NAME"))
            nurl = nurl & request.ServerVariables("SCRIPT_NAME")
            nurl = Left(nurl, InStrRev(nurl, "/"))'InStrRev()——倒数第一个"/"出现的正数位置数，Left(nurl, 正数位置数)=整个域名；
            filtpath = Replace(Str, nurl, "")'是过滤掉域名；
        Else'如果不是相对路径；
            filtpath = Str'（就用全域名！）
        End If
    End Function
    Public Function chkiplock()
        Dim IPlock
        IPlock = False
        Dim locklist
        locklist = Trim(setup(52, 0))
        If locklist = "" Then Exit Function
        Dim i, StrUserIP, StrKillIP
        StrUserIP = userip
        locklist = Split(locklist, vbCrLf)
        If StrUserIP = "" Then Exit Function
        StrUserIP = Split(userip, ".")
        If UBound(StrUserIP) <> 3 Then Exit Function
        For i = 0 To UBound(locklist)
            locklist(i) = Trim(locklist(i))
            If locklist(i) <> "" Then
                StrKillIP = Split(locklist(i), ".")
                If UBound(StrKillIP) <> 3 Then Exit For
                IPlock = True
                If (StrUserIP(0) <> StrKillIP(0)) And InStr(StrKillIP(0), "*") = 0 Then IPlock = False
                If (StrUserIP(1) <> StrKillIP(1)) And InStr(StrKillIP(1), "*") = 0 Then IPlock = False
                If (StrUserIP(2) <> StrKillIP(2)) And InStr(StrKillIP(2), "*") = 0 Then IPlock = False
                If (StrUserIP(3) <> StrKillIP(3)) And InStr(StrKillIP(3), "*") = 0 Then IPlock = False
                If IPlock Then Exit For
            End If
        Next
        chkiplock = IPlock
    End Function

    Public Function showpage(sfilename, totalnumber, maxperpage, ShowTotal, ShowAllPages, strUnit)'showtotal是否显示统计，unit计量单位；
        Dim n, i, strTemp, strUrl
        If totalnumber Mod maxperpage = 0 Then'嗯，计算出分页数n；
            n = totalnumber \ maxperpage
        Else
            n = totalnumber \ maxperpage + 1
        End If
        strTemp = "<div id=""showpage"">"
        If ShowTotal = True Then'看看是否允许显示 统计信息；
            strTemp = strTemp & "共" & totalnumber & strUnit & "&nbsp;&nbsp;"
        End If
        strUrl = JoinChar(sfilename)
        If CurrentPage < 2 Then'特殊处理第一页时；
                strTemp = strTemp & "首页 上一页&nbsp;"
        Else
                strTemp = strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
                strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>上一页</a>&nbsp;"
        End If
    
        If n - CurrentPage < 1 Then'特殊处理最后一页时；
                strTemp = strTemp & "下一页 尾页"
        Else
                strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>下一页</a>&nbsp;"
                strTemp = strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
        End If
        strTemp = strTemp & "&nbsp;页次：" & CurrentPage & "/" & n & "</strong>页 "
        strTemp = strTemp & "&nbsp;" & maxperpage & "" & strUnit & "/页"
        If ShowAllPages = True Then'是否显示下拉页码；
            strTemp = strTemp & "&nbsp;转到：<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"
            For i = 1 To n
                strTemp = strTemp & "<option value='" & i & "'"
                If CInt(CurrentPage) = CInt(i) Then strTemp = strTemp & " selected "
                strTemp = strTemp & ">" & i & "</option>"
            Next
            strTemp = strTemp & "</select>"
        End If
        strTemp = strTemp & "</div>"
        showpage = strTemp
    End Function
	
	Public Function Showpage_Front(sfilename, totalnumber, maxperpage, ShowTotal, ShowAllPages, strUnit)
        Dim n, i, strTemp, strUrl
        If totalnumber Mod maxperpage = 0 Then
            n = totalnumber \ maxperpage
        Else
            n = totalnumber \ maxperpage + 1
        End If
        strTemp = "<div id='showpage' style='text-align:center;'>"
        If ShowTotal = True Then
            strTemp = strTemp & "共" & totalnumber & strUnit & "&nbsp;&nbsp;"
        End If
        strUrl = JoinChar(sfilename)
		'strUrl = sfilename---------------wl-------------------
        If CurrentPage < 2 Then
                strTemp = strTemp & "首页 上一页&nbsp;"
        Else
                strTemp = strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
                strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>上一页</a>&nbsp;"
        End If
    
        If n - CurrentPage < 1 Then
                strTemp = strTemp & "下一页 尾页"
        Else
                strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>下一页</a>&nbsp;"
                strTemp = strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
        End If
        strTemp = strTemp & "&nbsp;页次：" & CurrentPage & "/" & n & "</strong>页 "
        strTemp = strTemp & "&nbsp;" & maxperpage & "" & strUnit & "/页"
        If ShowAllPages = True Then
            strTemp = strTemp & "&nbsp;转到：<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"
            For i = 1 To n
                strTemp = strTemp & "<option value='" & i & "'"
                If CInt(CurrentPage) = CInt(i) Then strTemp = strTemp & " selected "
                strTemp = strTemp & ">" & i & "</option>"
            Next
            strTemp = strTemp & "</select>"
        End If
        strTemp = strTemp & "</div>"
        Showpage_Front = strTemp
    End Function

    Public Function JoinChar(strUrl)
        If strUrl = "" Then
            JoinChar = ""
            Exit Function
        End If
        If InStr(strUrl, "?") < Len(strUrl) Then'如果?不在最后一个出现的话；
            If InStr(strUrl, "?") > 1 Then
                If InStr(strUrl, "&") < Len(strUrl) Then
                    JoinChar = strUrl & "&"
                Else
                    JoinChar = strUrl
                End If
            Else
                JoinChar = strUrl & "?"
            End If
        Else'如果?在最后一个出现，那么直接保留url；
            JoinChar = strUrl
        End If
    End Function
    Public Function htm2js(Str)
        If Str = "" Or IsNull(Str) Then Str = " "
        htm2js = "document.write('" & Replace(Replace(Replace(Replace(Str, "\", "\\"), "'", "\'"), vbCrLf, "\n"), Chr(13), "") & "');"
    End Function
    '将htm代码插入div
    Public Function htm2js_div(Str, divid)
        divid = Trim(divid)
        If Str = "" Or IsNull(Str) Then Str = " "
        htm2js_div = "if (chkdiv('" & divid & "')) {"
        htm2js_div = htm2js_div & "document.getElementById('"&divid&"')" & ".innerHTML='" & Replace(Replace(Replace(Replace(Str, "\", "\\"), "'", "\'"), vbCrLf, "\n"), Chr(13), "") & "';}"
        If divid = "subject" Then htm2js_div = htm2js_div & vbCrLf & "if (chkdiv('subject_l')) {document.getElementById('subject_l').innerHTML='" & Replace(Replace(Replace(Replace(Str, "\", "\\"), "'", "\'"), vbCrLf, "\n"), Chr(13), "") & "';}"
    End Function
    Public Function readfile(mPath, fName)
        On Error Resume Next
        Dim fs2, f2, fpath
        Set fs2 = server.CreateObject("Scripting.FileSystemObject")
        fpath = server.MapPath(mPath) & "\"
        fpath = fpath & fName
        Set f2 = fs2.OpenTextFile(fpath, 1, True)
        readfile = f2.ReadAll
        Set fs2 = Nothing
        Set f2 = Nothing
    End Function
    Public Function showsize(size)
        On Error Resume Next
        If size = "" Or IsNull(size) Then
            showsize = "0Byte"
            Exit Function
        End If
        showsize = size & "Byte"
        If size < 0 Then
            showsize = "0KB"
            Exit Function
        End If
        If size > 1024 Then'先检测是否大于第一个k（KB）；
           size = (size \ 1024)
           showsize = size & "KB"
        End If
        If size > 1024 Then'再检测是否大于第二个k（MB）；
           size = (size / 1024)
           showsize = FormatNumber(size, 2) & "MB"
        End If
        If size > 1024 Then'再检测是否大于第三个k（GB）；
           size = (size / 1024)
           showsize = FormatNumber(size, 2) & "GB"
        End If
    End Function
    Public Function ChkPost()
        Dim server_v1, server_v2
		ChkPost = False
		if true_domain=1 then
			ChkPost=true
			exit function
        end if
        server_v1 = CStr(request.ServerVariables("HTTP_REFERER"))'http://localhost:45233/test.asp
        server_v2 = CStr(request.ServerVariables("SERVER_NAME"))'localhost	(www.myhomestay.com.cn)
        If Mid(server_v1, 8, Len(server_v2)) = server_v2 Then ChkPost = True'截取字符数Len(server_v2)；
		ChkPost = True'''WL;强制避免限制；
    End Function
    Public Function filt_badstr(sSql)'防注入功能WL[2007-09-23];处理数字参数；
		 If ISNull(sSql) Then Exit Function
		 sSql=Trim(sSql)
		 If sSql="" Then Exit Function 
		 sSql=Replace(sSql,Chr(0),"") 
		 sSql=Replace(sSql,"'","''")
		 
		 sSql=Replace(sSql,"%","％")
		 sSql=Replace(sSql,"-","－")
		 sSql=Replace(sSql,"(","（")  'WL
		 sSql=Replace(sSql,")","）")  'WL
		 sSql=Replace(sSql,";","；")  'WL
		 sSql=Replace(sSql,"#","＃")  'WL
		 filt_badstr=sSql
	End Function
    Public Function filt_astr(Str, n)'过滤字符串，并按照数量截取字符串；
        If IsNull(Str) Then
            filt_astr = ""
            Exit Function
        End If
        filt_astr = filt_badword(Str)
        filt_astr = InterceptStr(filt_astr, n)
    End Function
    Public Function filt_html(Str)
        On Error Resume Next
        If Str = "" Then
            filt_html = ""
        Else
            Str = Replace(Str, ">", "&gt;")
            Str = Replace(Str, "<", "&lt;")
            Str = Replace(Str, Chr(32), "&nbsp;")
            Str = Replace(Str, Chr(9), "&nbsp;")
            Str = Replace(Str, Chr(34), "&quot;")
            Str = Replace(Str, Chr(39), "&#39;")
            Str = Replace(Str, Chr(13), "")
            Str = Replace(Str, Chr(10) & Chr(10), "&nbsp; ")
            Str = Replace(Str, Chr(10), "&nbsp; ")
            filt_html = Str
        End If
    End Function
    Public Function filt_html_b(fString)
        If Not IsNull(fString) Then
            fString = Replace(fString, ">", "&gt;")
            fString = Replace(fString, "<", "&lt;")
            fString = Replace(fString, Chr(32), " ")
            fString = Replace(fString, Chr(9), " ")
            fString = Replace(fString, Chr(34), "&quot;")
            'fString = Replace(fString, CHR(39), "&#39;")
            fString = Replace(fString, Chr(13), "")
            fString = Replace(fString, Chr(10) & Chr(10), "</p><p> ")
            fString = Replace(fString, Chr(10), "<br> ")
            filt_html_b = fString
        End If
    End Function
	
	'功能描述： 用正则除去HTML标记
	'输入参数： fString : 待处理的字符串
	'返回值： String : 已处理的字符串
''	Public Function RemoveHTMLTag(fString)
''	 Dim ObjReg
''	 Set ObjReg = New Regexp
''	 ObjReg.IgnoreCase = True
''	 ObjReg.Global = True
''
''
''	 ObjReg.Pattern = "<[^>]+>|</[^>]+>"
''	 fString = ObjReg.Replace(fString,"")
''	 RemoveHTMLTag = fString
''	 
''	 Set ObjReg = Nothing
''	End Function 
''	'自动去除字符中含有html代码的几个ASP函数
''	Public Function Replacehtml(Textstr)
''		Dim Str,re
''		Str=Textstr
''		Set re=new RegExp
''		re.IgnoreCase =True
''		re.Global=True
''		re.Pattern="<(.[^>]*)>"
''		Str=re.Replace(Str, "")
''		Set Re=Nothing
''		Replacehtml=Str
''	End Function
	Public function nohtml(str) '要的就是你WL;
		dim re 
		Set re=new RegExp 
		re.IgnoreCase =true 
		re.Global=True 
		re.Pattern="(\<.[^\<]*\>)" 
		str=re.replace(str," ") 
		re.Pattern="(\<\/[^\<]*\>)" 
		str=re.replace(str," ") 
		nohtml=str 
		set re=nothing 
	end function
	
	Public function WeekDate(strTime)
	dim xingqi,xingqiji
	   xingqi=weekday(strTime)
	   select case xingqi
	   case "1"
	   xingqiji="日"
	   case "2"
	   xingqiji="一"
	   case "3"
	   xingqiji="二"
	   case "4"
	   xingqiji="三"
	   case "5"
	   xingqiji="四"
	   case "6"
	   xingqiji="五"
	   case "7"
	   xingqiji="六"
	   end select
	   WeekDate = xingqiji
	end function
	
	Public function WeekDateE(strTime)
	dim xingqi,xingqiji
	   xingqi=weekday(strTime)
	   select case xingqi
	   case "1"
	   xingqiji="Sunday"
	   case "2"
	   xingqiji="Monday"
	   case "3"
	   xingqiji="Tuesday"
	   case "4"
	   xingqiji="Wednesday"
	   case "5"
	   xingqiji="Thursday"
	   case "6"
	   xingqiji="Friday"
	   case "7"
	   xingqiji="Saturday"
	   end select
	   WeekDateE = xingqiji
	end function

Public Function y()'下拉显示年；
dim option_y
	option_y=option_y&"<option value=''>选择</option>"
	option_y=option_y&"<option value='1900'>1900</option>"
	option_y=option_y&"<option value='1901'>1901</option>"
	option_y=option_y&"<option value='1902'>1902</option>"
	option_y=option_y&"<option value='1903'>1903</option>"
	option_y=option_y&"<option value='1904'>1904</option>"
	option_y=option_y&"<option value='1905'>1905</option>"
	option_y=option_y&"<option value='1906'>1906</option>"
	option_y=option_y&"<option value='1907'>1907</option>"
	option_y=option_y&"<option value='1908'>1908</option>"
	option_y=option_y&"<option value='1909'>1909</option>"
	option_y=option_y&"<option value='1910'>1910</option>"
	option_y=option_y&"<option value='1911'>1911</option>"
	option_y=option_y&"<option value='1912'>1912</option>"
	option_y=option_y&"<option value='1913'>1913</option>"
	option_y=option_y&"<option value='1914'>1914</option>"
	option_y=option_y&"<option value='1915'>1915</option>"
	option_y=option_y&"<option value='1916'>1916</option>"
	option_y=option_y&"<option value='1917'>1917</option>"
	option_y=option_y&"<option value='1918'>1918</option>"
	option_y=option_y&"<option value='1919'>1919</option>"
	option_y=option_y&"<option value='1920'>1920</option>"
	option_y=option_y&"<option value='1921'>1921</option>"
	option_y=option_y&"<option value='1922'>1922</option>"
	option_y=option_y&"<option value='1923'>1923</option>"
	option_y=option_y&"<option value='1924'>1924</option>"
	option_y=option_y&"<option value='1925'>1925</option>"
	option_y=option_y&"<option value='1926'>1926</option>"
	option_y=option_y&"<option value='1927'>1927</option>"
	option_y=option_y&"<option value='1928'>1928</option>"
	option_y=option_y&"<option value='1929'>1929</option>"
	option_y=option_y&"<option value='1930'>1930</option>"
	option_y=option_y&"<option value='1931'>1931</option>"
	option_y=option_y&"<option value='1932'>1932</option>"
	option_y=option_y&"<option value='1933'>1933</option>"
	option_y=option_y&"<option value='1934'>1934</option>"
	option_y=option_y&"<option value='1935'>1935</option>"
	option_y=option_y&"<option value='1936'>1936</option>"
	option_y=option_y&"<option value='1937'>1937</option>"
	option_y=option_y&"<option value='1938'>1938</option>"
	option_y=option_y&"<option value='1939'>1939</option>"
	option_y=option_y&"<option value='1940'>1940</option>"
	option_y=option_y&"<option value='1941'>1941</option>"
	option_y=option_y&"<option value='1942'>1942</option>"
	option_y=option_y&"<option value='1943'>1943</option>"
	option_y=option_y&"<option value='1944'>1944</option>"
	option_y=option_y&"<option value='1945'>1945</option>"
	option_y=option_y&"<option value='1946'>1946</option>"
	option_y=option_y&"<option value='1947'>1947</option>"
	option_y=option_y&"<option value='1948'>1948</option>"
	option_y=option_y&"<option value='1949'>1949</option>"
	option_y=option_y&"<option value='1950'>1950</option>"
	option_y=option_y&"<option value='1951'>1951</option>"
	option_y=option_y&"<option value='1952'>1952</option>"
	option_y=option_y&"<option value='1953'>1953</option>"
	option_y=option_y&"<option value='1954'>1954</option>"
	option_y=option_y&"<option value='1955'>1955</option>"
	option_y=option_y&"<option value='1956'>1956</option>"
	option_y=option_y&"<option value='1957'>1957</option>"
	option_y=option_y&"<option value='1958'>1958</option>"
	option_y=option_y&"<option value='1959'>1959</option>"
	option_y=option_y&"<option value='1960'>1960</option>"
	option_y=option_y&"<option value='1961'>1961</option>"
	option_y=option_y&"<option value='1962'>1962</option>"
	option_y=option_y&"<option value='1963'>1963</option>"
	option_y=option_y&"<option value='1964'>1964</option>"
	option_y=option_y&"<option value='1965'>1965</option>"
	option_y=option_y&"<option value='1966'>1966</option>"
	option_y=option_y&"<option value='1967'>1967</option>"
	option_y=option_y&"<option value='1968'>1968</option>"
	option_y=option_y&"<option value='1969'>1969</option>"
	option_y=option_y&"<option value='1970' selected>1970</option>"
	option_y=option_y&"<option value='1971'>1971</option>"
	option_y=option_y&"<option value='1972'>1972</option>"
	option_y=option_y&"<option value='1973'>1973</option>"
	option_y=option_y&"<option value='1974'>1974</option>"
	option_y=option_y&"<option value='1975'>1975</option>"
	option_y=option_y&"<option value='1976'>1976</option>"
	option_y=option_y&"<option value='1977'>1977</option>"
	option_y=option_y&"<option value='1978'>1978</option>"
	option_y=option_y&"<option value='1979'>1979</option>"
	option_y=option_y&"<option value='1980'>1980</option>"
	option_y=option_y&"<option value='1981'>1981</option>"
	option_y=option_y&"<option value='1982'>1982</option>"
	option_y=option_y&"<option value='1983'>1983</option>"
	option_y=option_y&"<option value='1984'>1984</option>"
	option_y=option_y&"<option value='1985'>1985</option>"
	option_y=option_y&"<option value='1986'>1986</option>"
	option_y=option_y&"<option value='1987'>1987</option>"
	option_y=option_y&"<option value='1988'>1988</option>"
	option_y=option_y&"<option value='1989'>1989</option>"
	option_y=option_y&"<option value='1990'>1990</option>"
	option_y=option_y&"<option value='1991'>1991</option>"
	option_y=option_y&"<option value='1992'>1992</option>"
	option_y=option_y&"<option value='1993'>1993</option>"
	option_y=option_y&"<option value='1994'>1994</option>"
	option_y=option_y&"<option value='1995'>1995</option>"
	option_y=option_y&"<option value='1996'>1996</option>"
	option_y=option_y&"<option value='1997'>1997</option>"
	option_y=option_y&"<option value='1998'>1998</option>"
	option_y=option_y&"<option value='1999'>1999</option>"
	option_y=option_y&"<option value='2000'>2000</option>"
	option_y=option_y&"<option value='2001'>2001</option>"
	option_y=option_y&"<option value='2002'>2002</option>"
	option_y=option_y&"<option value='2003'>2003</option>"
	option_y=option_y&"<option value='2004'>2004</option>"
	option_y=option_y&"<option value='2005'>2005</option>"
	option_y=option_y&"<option value='2006'>2006</option>"
	option_y=option_y&"<option value='2006'>2007</option>"
	option_y=option_y&"<option value='2006'>2008</option>"
	y = option_y
End Function

Public Function m()'下拉显示月；
dim option_m
	option_m=option_m&"<option value=''>选择</option>"
	option_m=option_m&"<option value='1'>1</option>"
	option_m=option_m&"<option value='2'>2</option>"
	option_m=option_m&"<option value='3'>3</option>"
	option_m=option_m&"<option value='4'>4</option>"
	option_m=option_m&"<option value='5'>5</option>"
	option_m=option_m&"<option value='6'>6</option>"
	option_m=option_m&"<option value='7'>7</option>"
	option_m=option_m&"<option value='8'>8</option>"
	option_m=option_m&"<option value='9'>9</option>"
	option_m=option_m&"<option value='10'>10</option>"
	option_m=option_m&"<option value='11'>11</option>"
	option_m=option_m&"<option value='12'>12</option>"
	m = option_m
End Function

Public Function d()'下拉显示日；
dim option_d
	option_d=option_d&"<option value=''>选择</option>"
	option_d=option_d&"<option value='1'>1</option>"
	option_d=option_d&"<option value='2'>2</option>"
	option_d=option_d&"<option value='3'>3</option>"
	option_d=option_d&"<option value='4'>4</option>"
	option_d=option_d&"<option value='5'>5</option>"
	option_d=option_d&"<option value='6'>6</option>"
	option_d=option_d&"<option value='7'>7</option>"
	option_d=option_d&"<option value='8'>8</option>"
	option_d=option_d&"<option value='9'>9</option>"
	option_d=option_d&"<option value='10'>10</option>"
	option_d=option_d&"<option value='11'>11</option>"
	option_d=option_d&"<option value='12'>12</option>"
	option_d=option_d&"<option value='13'>13</option>"
	option_d=option_d&"<option value='14'>14</option>"
	option_d=option_d&"<option value='15'>15</option>"
	option_d=option_d&"<option value='16'>16</option>"
	option_d=option_d&"<option value='17'>17</option>"
	option_d=option_d&"<option value='18'>18</option>"
	option_d=option_d&"<option value='19'>19</option>"
	option_d=option_d&"<option value='20'>20</option>"
	option_d=option_d&"<option value='21'>21</option>"
	option_d=option_d&"<option value='22'>22</option>"
	option_d=option_d&"<option value='23'>23</option>"
	option_d=option_d&"<option value='24'>24</option>"
	option_d=option_d&"<option value='25'>25</option>"
	option_d=option_d&"<option value='26'>26</option>"
	option_d=option_d&"<option value='27'>27</option>"
	option_d=option_d&"<option value='28'>28</option>"
	option_d=option_d&"<option value='29'>29</option>"
	option_d=option_d&"<option value='30'>30</option>"
	option_d=option_d&"<option value='31'>31</option>"
	d = option_d
End Function





    Public Function strLength(Str)
        On Error Resume Next
        Dim WINNT_CHINESE
        WINNT_CHINESE = (Len("中国") = 2)
        If WINNT_CHINESE Then
            Dim l, t, c
            Dim i
            l = Len(Str)
            t = l
            For i = 1 To l
                c = Asc(Mid(Str, i, 1))
                If c < 0 Then c = c + 65536
                If c > 255 Then
                    t = t + 1
                End If
            Next
            strLength = t
        Else
        strLength = Len(Str)
        End If
        If Err.Number <> 0 Then Err.Clear
    End Function
    Public Function InterceptStr(txt, length)'截取函数~用于计数 中文和英文字符的混合文字占用的字符串数 来截取；
        Dim x, y, ii
        txt = Trim(txt)
        x = Len(txt)
        y = 0
        If x >= 1 Then
            For ii = 1 To x
                If Asc(Mid(txt, ii, 1)) < 0 Or Asc(Mid(txt, ii, 1)) > 255 Then '如果是汉字
                    y = y + 2
                Else
                    y = y + 1
                End If
                If y >= length Then
                    txt = Left(Trim(txt), ii) '字符串限长
                    Exit For
                End If
            Next
            InterceptStr = txt
        Else
            InterceptStr = ""
        End If
    End Function
    '读取用户目录对应绑定的路径，未绑定返回空
    Public Function getdirdomain(udir)
        Dim tmp1, tmp2, Str
        Str = Application(cache_name & "dirdomain")
        udir = Trim(udir)
        tmp1 = InStr(Str, udir & "!!??((")
        tmp2 = Len(udir & "!!??((") + tmp1
        If tmp1 > 0 Then
            getdirdomain = Mid(Str, tmp2, InStr(tmp1, Str, "##))==") - tmp2)
        Else
            getdirdomain = ""
    End If
    End Function
    Public Function GetUrl()
        On Error Resume Next
        Dim strTemp
        If LCase(request.ServerVariables("HTTPS")) = "off" Then
        strTemp = "http://"
        Else
        strTemp = "https://"
        End If
        strTemp = strTemp & request.ServerVariables("SERVER_NAME")
        If request.ServerVariables("SERVER_PORT") <> 80 Then strTemp = strTemp & ":" & request.ServerVariables("SERVER_PORT")
        strTemp = strTemp & request.ServerVariables("URL")
        If Trim(request.QueryString) <> "" Then strTemp = strTemp & "?" & Trim(request.QueryString)
        GetUrl = strTemp
    End Function
    Public Function trueurl(strContent)
        On Error Resume Next
        Dim tempReg, url
        url = Trim("http://" & request.ServerVariables("SERVER_NAME"))
        url = LCase(url & request.ServerVariables("SCRIPT_NAME"))
        url = Left(url, InStrRev(url, "/"))
        Set tempReg = New RegExp
        tempReg.IgnoreCase = True
        tempReg.Global = True
        tempReg.Pattern = "(^.*\/).*$" '含文件名的标准路径
        url = tempReg.Replace(url, "$1")
        tempReg.Pattern = "((?:src|href).*?=[\'\u0022](?!ftp|http|https|mailto))"
        trueurl = tempReg.Replace(strContent, "$1" + url)
        Set tempReg = Nothing
    End Function
    Public Function IsValidEmail(email)'有效的e-mail格式；
        Dim names, name, i, c
        IsValidEmail = True
        names = Split(email, "@")
        If UBound(names) <> 1 Then
           IsValidEmail = False
           Exit Function
        End If
        For Each name In names
           If Len(name) <= 0 Then
             IsValidEmail = False
             Exit Function
           End If
           For i = 1 To Len(name)'截取每一个字符，并对其操作（检测）；
             c = LCase(Mid(name, i, 1))
             If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
               IsValidEmail = False
               Exit Function
             End If
           Next
           If Left(name, 1) = "." Or Right(name, 1) = "." Then
              IsValidEmail = False
              Exit Function
           End If
        Next
        If InStr(names(1), ".") <= 0 Then
           IsValidEmail = False
           Exit Function
        End If
        i = Len(names(1)) - InStrRev(names(1), ".")
        If i <> 2 And i <> 3 Then
           IsValidEmail = False
           Exit Function
        End If
        If InStr(email, "..") > 0 Then
           IsValidEmail = False
        End If
    End Function
    Public Function chkdomain(domain)
        Dim name, i, c
        name = domain
        chkdomain = True
        If Len(name) <= 0 Then
            chkdomain = False
            Exit Function
        End If
        For i = 1 To Len(name)
           c = LCase(Mid(name, i, 1))
            If InStr("abcdefghijklmnopqrstuvwxyz-", c) <= 0 And Not IsNumeric(c) Then
               chkdomain = False
            Exit Function
           End If
       Next
    End Function
    Public Function CodeCookie(Str)
    If is_password_cookies = 1 Then
        Dim i
        Dim StrRtn
        For i = Len(Str) To 1 Step -1
            StrRtn = StrRtn & AscW(Mid(Str, i, 1))
            If (i <> 1) Then StrRtn = StrRtn & "a"
        Next
        CodeCookie = StrRtn
    Else
        CodeCookie = Str
    End If
    End Function
    
    Public Function DecodeCookie(Str)
    If is_password_cookies = 1 Then
        Dim i
        Dim StrArr, StrRtn
        StrArr = Split(Str, "a")
        For i = 0 To UBound(StrArr)
            If IsNumeric(StrArr(i)) = True Then
                StrRtn = ChrW(StrArr(i)) & StrRtn
            Else
                StrRtn = Str
                Exit Function
            End If
        Next
        DecodeCookie = StrRtn
    Else
        DecodeCookie = Str
    End If
    End Function
    Private Sub class_terminate()
    On Error Resume Next
        If IsObject(conn) Then conn.Close: Set conn = Nothing
    End Sub
    
    Public Function BuildFile(ByVal sFile, ByVal sContent)
        Dim oFSO, oStream
        If is_gb2312 = 1 Then
            Set oFSO = server.CreateObject("Scripting.FileSystemObject")
            'Response.Write "目录1：" & sFile & "<br>"
            Set oStream = oFSO.CreateTextFile(sFile, True)
            oStream.Write sContent
            oStream.Close
            Set oStream = Nothing
            Set oFSO = Nothing
        Else
            Set oStream = server.CreateObject("ADODB.Stream")
            With oStream
                .Type = 2
                .Mode = 3
                .Open
                '.Charset = "utf-8"
                .Charset = "gb2312"
                .Position = oStream.size
                .WriteText = sContent
                .SaveToFile sFile, 2
                .Close
            End With
            Set oStream = Nothing
        End If
    End Function
    '过滤掉flash UBB标记
    '[flash=500,350]http://www.kunfu.net/movie.swf[/flash]
    Function FilterUBBFlash(ByVal strFlash)
        Dim strFlash1
        strFlash1 = LCase(strFlash)
        If InStr(strFlash1, "[/flash]") > 0 Then
            strFlash1 = Replace(strFlash1, "[/flash]", "[ /flash ]")
            strFlash1 = Replace(strFlash1, "[flash", "[ flash ")
            FilterUBBFlash = strFlash1
        Else
            FilterUBBFlash = strFlash
        End If
    End Function
	
	
	'	例子
	'TypeName = requestID_selectName("clinic_Type","classname","Where clinicTypeID="&request("TypeNumber"))
					'有3个参数，分别是： 
					'表名；
					'ID名；
					'显示名；
					'加上‘Where’条件语句；
	'例子 — — If Bed_ID<>0 Then Bed_Name = requestID_selectName("reg_Bed","classname","Where classid="&Bed_ID)		
	'获取 [某ID] 所对应的 [某字段值]~
	FUNCTION requestID_selectName(Table,Field_name,Where)
	   '有3个参数，分别是：
		'表名；
		'ID名；
		'显示名；
		'加上'Where'条件语句；
	'例子 — --- If Bed_ID<>0 Then Bed_Name = requestID_selectName("reg_Bed","classname","Where classid="&Bed_ID)
	dim scID_Rs_Temp
	set scID_Rs_Temp=conn.execute("select "&Field_name&" from "&Table&" "&Where)  
	 If scID_Rs_Temp.eof Then
	  requestID_selectName = "没有此记录"
	 Else
	  requestID_selectName = scID_Rs_Temp(Field_name)
	 
	 End If
	 
	scID_Rs_Temp.close
	
	
	End FUNCTION


End Class
%>




