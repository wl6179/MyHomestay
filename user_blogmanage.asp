<!--#include file="user_top.asp"-->
<!--#include file="inc/class_blog.asp"-->
<div id="head">
	<div id="head2">
	
		<div id="head2-logo">
			
			<ul>
				<li class="active"><a href="http://www.myhomestay.com.cn">��������</a></li>
				<li><a href="http://www.myhomestay.com.cn">English</a></li>
				<!--<li><a href="http://www.myhomestay.com.cn">Janpan</a></li>-->
			</ul>
			
		</div>
		
		
		<div id="head2-menu">
			<div id="head2-search">
				<span>�ͻ���ѯ����</span>��<span>010-85493388</span>/<span>13910652850</span> &nbsp;&nbsp;&nbsp;
				<!--<a href="http://www.myhomestay.com.cn">��Ϊ��ҳ</a>&nbsp;&nbsp;&nbsp;-->
				<!--<a href="#" onClick="javascript:AddFavoriteOnClick();">���ո�������ղؼ�</a>-->&nbsp;
				<!--<a href="http://www.myhomestay.com.cn">��������</a>&nbsp;&nbsp;&nbsp;-->
			</div>
			
			<div id="divCN-EN">
			<ul>
				<li><a href="/index.asp">��ҳ<br />Home</a></li><!--#include file="menu/incIndexNav.htm"-->
	
				
			</ul>
			</div>
		</div>
	
	</div>
	
</div>
	
	
	<div id="body"><!--#body-->
	<div id="gbody"><!--#gbody-->
	
		<div id="left-margin"></div>
		
		<div id="main"><!--the main content area-->
			
			
			
			
			<div id="mainContentList">
				<div id="main-top-round">
					<div id="main-top-round-left"></div>
					<div id="main-top-round-right"></div>
				</div>
				
				<div id="main-main-round">
					<style>
						table {
						background:#ffffff;
						border:1px solid #999999;
						border-collapse: collapse;
						font-size:13px;
						}
						td {
						padding:0px 0px 0px 12px;
						
						height:18px;
						border:1px solid #999999;
						}
							td.tr02-1 {
							width:30px;
							}
						 	td.tr02-2 {
							width:60px;
							}
							td.tr02-3 {
							width:180px;
							}
							td.tr02-4 {
							width:60px;
							}
							td.tr02-5 {
							width:80px;
							}
							td.tr02-6 {
							width:40px;
							}
							td.tr02-7 {
							width:120px;
							}
							
						.tr01 {
						/*font:bold;*/
						background:#A1E5FA;
						font-boil:2px;
						color:#4f6b72;
						}
						.tr02 {
						background:#ffffff;
						color:#797268;
						}
						.tr03 {
						background:#F7FDFB;
						color:#4f6b72;
						}
					</style>
					<div id="main-TextArea01">
					
					<% dim temp_word
					usersearch=5
					If Request("action")="downlog" Then
						temp_word="����"
					ElseIf Request("usersearch")=5 Then
						temp_word="�ݸ���"
					Else
						temp_word="����"
					End If
					
					%>
					
					
						<center><h2>MyHomestay <%=temp_word%><%=tName%></h2></center>
						


    <div class="content_top">
            <div class="side_d1 side11"></div>
            <div class="side_d2 side21"></div>
           
    </div>
  
<%
Const MaxPerPage = 20
Dim strFileName
Dim totalPut, TotalPages
Dim rs, sql
Dim id, usersearch, Keyword, strField, uid, action
Dim selectsub, substr, wsql, allsub
Keyword = Trim(request("keyword"))
If Keyword <> "" Then Keyword = oblog.filt_badstr(Keyword)
strField = Trim(request("Field"))
usersearch = Trim(request("usersearch"))
selectsub = Trim(request("selectsub"))
action = Trim(request("action"))
id = oblog.filt_badstr(Trim(request("id")))
uid = oblog.filt_badstr(Trim(request("uid")))
strFileName = "user_blogmanage.asp?t=" & t
If usersearch = "" Then
    usersearch = 0
Else
    usersearch = CLng(usersearch)
    strFileName = "user_blogmanage.asp?usersearch=" & usersearch & "&t=" & t
End If
If selectsub <> "" Then
    selectsub = CLng(selectsub)
    strFileName = "user_blogmanage.asp?usersearch=10&selectsub=" & selectsub & "&t=" & t
Else
    selectsub = 0
End If
If Keyword <> "" Then
    strFileName="user_blogmanage.asp?usersearch=10&keyword="&Keyword&"&Field="&strField & "&t="  & t
End If
If request("page") <> "" Then CurrentPage = CInt(request("page")) Else CurrentPage = 1
if oblog.logined_ulevel<>9 then wsql=" and ( userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&" )"

Select Case action
    Case "del"
    	Call delblog
    Case "Lock"
    	Call locklog
    Case "UnLock"
    	Call passlog
    Case "isbest"
    	Call isbest
    Case "unisbest"
    	Call unisbest
    Case "Move"
    	Call moveblog
    Case "updatelog"
    	Call updatelog
    Case "downlog"
    	Call downlog
    Case Else
	    Call top
	    Call main
End Select
Set rs = Nothing
%>


  
  
  
  </div>
  
  					<!--wl-->
					</div>
					
					
					<div id="main-main-round2">
						<center>
					  
					    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="585" height="90">
                          <param name="movie" value="images/90620_hp585_90.swf" />
                          <param name="quality" value="high" />
                          <embed src="images/90620_hp585_90.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="585" height="90"></embed>
					      </object>
					  	</center>
					</div>
					
					
		
					
					
				
					<div id="main-banner01">
						<div id="main-bottom-round-left"></div>
						<div id="main-bottom-round-right"></div>
					</div>
				</div>
				
				
				<div id="publicList">
					<div id="public-top-round">
						<div id="public-top-round-left"></div>
						<div id="public-top-round-right"></div>
					</div>
					
					<div id="public-main-round">
					<%'main()%>
						<div class="public-TextArea01">
							<p class="public-title01">
								��Ŀ��������
							</p>
							
							<!--#include file="menu/inc-menu.asp"-->
							
						</div>
						
					
					</div>
				
					<div id="public-bottom-round">
						<div id="public-bottom-round-left"></div>
						<div id="public-bottom-round-right"></div>
					</div>
				</div>				
			</div>
			

</div>
		</div><!--#gbody-->	
	  </div><!--#body-->


  

	  		
<%=oblog.user_copyright%>
			
			
	  	


  <script src="inc/divCN-EN.js" type="text/javascript"></script>
  </body>
</html>
<%
Sub top()
%>
<div id="list" style="width:99%;text-align:left;">
    <h2 style="width:100%;font-size:13px; text-align:left;padding:8px;">
    <form name="form1" action="user_blogmanage.asp" method="get">
    	
		<input type="hidden" name="t" value="<%=t%>">

���ٲ���:
        <select size=1 name="usersearch" onChange="javascript:submit()">
          <option value="10" selected>��ѡ���������</option>
          <option value="0">�г�����<%=tName%></option>
          <option value="1">δͨ����˵�<%=tName%></option>
          <option value="2">��ͨ����˵�<%=tName%></option>
          <option value="3">�Ƽ�<%=tName%></option>
           <option value="5">�ݸ���</option>
          <%if oblog.logined_ulevel=9 then%>
          <option value="4">�ҵ�<%=tName%></option>
          <%end if%>
        </select> &nbsp;
      
	  ��ר�����:&nbsp;
        <select name="selectsub" id="selectsub" onChange="javascript:submit()">
<%
response.write "<option value=''>��ѡ��ר��</option>"
Set rs = oblog.execute("select subjectid,subjectname from oblog_subject where userid=" & oblog.logined_uid & " And subjecttype=" & t)
While Not rs.EOF
    substr=substr&"<option value="&rs(0)&">"&rs(1)&"</option>"
    allsub=allsub&rs(0)&"!!??(("&rs(1)&"##))=="
    rs.movenext
Wend

response.write (substr)
response.write "<option value=0>δ����</option>"
%>
        </select>&nbsp;
<br /><%=tName%>����:
        <select name="Field" id="Field">
            <%if oblog.logined_ulevel=9 then%>
            <option value="userid">�û�id</option>
            <option value="user">�û���</option>
            <%end if%>
            <option value="id"><%=tName%>ID��</option>
            <option value="topic" selected><%=tName%>����</option>
         </select>
		 
         <input name="Keyword" type="text" id="Keyword" size="20" maxlength="30">
		 
         <input type="submit" name="Submit2" value=" ���� ">
  </form>
    </h2>
<%
End Sub

Sub main()
%>
<style type="text/css">
<!-- 
    #list ul{ width: 95%;;}
    #list ul li.t1 { width: 50px}
    #list ul li.t2 { width: 90px}
    #list ul li.t3 { width: 190px}
    #list ul li.t4 { width: 90px}
    #list ul li.t5 { width: 100px}
    #list ul li.t6 { width: 60px}
    #list ul li.t7 { width: 50px}
    #list ul li.t8 { width: 60px}
    #list ul li.t9 { width: 90px}
-->
</style>

<style type="text/css">
<!-- 
	ul.list_top{ width: 100%;padding:0; font-size:12px;}
	ul.list_top li.t1 { width: 50px;float:left;}
	ul.list_top li.t2 { width: 90px;float:left;}
	ul.list_top li.t3 { width: 210px;float:left;}
	ul.list_top li.t4 { width: 90px;float:left;}
	ul.list_top li.t5 { width: 100px;float:left;}
	ul.list_top li.t6 { width: 80px;float:left;}
	ul.list_top li.t7 { width: 70px;float:left;}
	ul.list_top li.t8 { width: 80px;float:left;}
	ul.list_top li.t9 { width: 110px;float:left;}
	 
     .list_content{ width: 100%;padding:0px; font-size:12px;}
     .list_content li.t1 { width: 50px;float:left;}
     .list_content li.t2 { width: 90px;float:left;}
     .list_content li.t3 { width: 210px;float:left;}
     .list_content li.t4 { width: 90px;float:left;}
     .list_content li.t5 { width: 100px;float:left;}
     .list_content li.t6 { width: 80px;float:left;}
     .list_content li.t7 { width: 70px;float:left;}
     .list_content li.t8 { width: 80px;float:left;}
     .list_content li.t9 { width: 110px;float:left;}
-->
</style>


<%
    Dim strGuide, selectsql
    selectsql = "logid,userid,iis,commentnum,topic,author,addtime,logfile,isbest,isdraft,passcheck,issave,subjectid,classid"
    strGuide = "<h1 style='font-size:13px;'>��ǰѡ��&nbsp;&gt;&gt;&nbsp;"
    Select Case usersearch
        Case 0
            If oblog.logined_ulevel = 9 Then
                sql="select top 1000 "&selectsql&" from [oblog_log] Where logtype=" & t & " order by logid desc"
            Else
                sql="select "&selectsql&" from oblog_log where logtype=" & t & " And ( userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&") order by addtime desc"
            End If
            strGuide = strGuide & "����" & tName
        Case 1
            If oblog.logined_ulevel = 9 Then
                sql="select "&selectsql&" from [oblog_log] where passcheck=0 And logtype=" & t & " order by logid desc"
            Else
                sql="select "&selectsql&" from [oblog_log] where passcheck=0 And logtype=" & t & " and ( userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&" ) order by addtime desc"
            End If
            strGuide = strGuide & "δͨ�����" & tName
        Case 2
            If oblog.logined_ulevel = 9 Then
                sql="select "&selectsql&" from [oblog_log] where passcheck=1  And logtype=" & t & " order by logid desc"
            Else
                sql="select "&selectsql&" from [oblog_log] where passcheck=1  And logtype=" & t & " and ( userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&" ) order by addtime desc"
            End If
            strGuide = strGuide & "��ͨ�����" & tName
        Case 3
            If oblog.logined_ulevel = 9 Then
                sql="select "&selectsql&" from [oblog_log] where isbest=1  And logtype=" & t & " order by logid desc"
            Else
                sql="select "&selectsql&" from [oblog_log] where isbest=1  And logtype=" & t & " and ( userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&" ) order by addtime desc"
            End If
            strGuide = strGuide & "�Ƽ�" & tName
        Case 4
            sql="select "&selectsql&" from [oblog_log] where ( userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&" )  And logtype=" & t & " order by addtime desc"
            strGuide = strGuide & "�ҵ�" & tName
        Case 5
            If oblog.logined_ulevel = 9 Then
                sql="select "&selectsql&" from [oblog_log] where isdraft=1  And logtype=" & t & " order by logid desc"
            Else
                sql="select "&selectsql&" from [oblog_log] where isdraft=1  And logtype=" & t & "  and ( userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&" ) order by addtime desc"
            End If
            strGuide = strGuide & "�ݸ���"
        Case 10
            If Keyword = "" Then
                sql="select "&selectsql&" from [oblog_log] where (userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&") and subjectid="&selectsub&"   And logtype=" & t & " order by addtime desc"
                strGuide=strGuide & "ר��idΪ"&selectsub&"��" & tName
            Else
                Select Case strField
                Case "userid"
                If IsNumeric(Keyword) = False Then
                        oblog.adderrstr ("�û�id������������")
                        oblog.showusererr
                    Else
                        If oblog.logined_ulevel = 9 Then
                        sql="select "&selectsql&" from [oblog_log] where userid = "&Keyword&"   And logtype=" & t & " order by logid desc"
                        strGuide = strGuide & "�û�idΪ<font color=red> " & Keyword & " </font>��" & tName
                    End If
                    End If
                Case "user"
                        If oblog.logined_ulevel = 9 Then
                        sql="select "&selectsql&" from [oblog_log] where author = '"&Keyword&"'   And logtype=" & t & " order by logid desc"
                        strGuide = strGuide & "�û���Ϊ<font color=red> " & Keyword & " </font>��" & tName
                    End If
                Case "id"
                    If IsNumeric(Keyword) = False Then
                        oblog.adderrstr (tName & "id������������")
                        oblog.showusererr
                    Else
                        If oblog.logined_ulevel = 9 Then
                            sql="select "&selectsql&" from [oblog_log] where logid =" & Clng(Keyword)  & " And logtype=" & t
                        Else
                            sql="select "&selectsql&" from [oblog_log] where logid =" & Clng(Keyword)&"  And logtype=" & t & " and (userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&")"
                        End If
                        strGuide = strGuide & "id����<font color=red> " & CLng(Keyword) & " </font>��" & tName
                    End If
                Case "topic"
                    If oblog.logined_ulevel = 9 Then
                        sql="select top 500 "&selectsql&" from [oblog_log] where topic like '%" & Keyword & "%'   And logtype=" & t & " order by logid desc"
                    Else
                        sql="select "&selectsql&" from [oblog_log] where topic like '%" & Keyword & "%' and ( userid="&oblog.logined_uid&" or authorid="&oblog.logined_uid&" )   And logtype=" & t & " order by addtime desc"
                    End If
                    strGuide = strGuide & "�����к��С� <font color=red>" & Keyword & "</font> ����" & tName
                End Select
            End If
        Case Else
            oblog.adderrstr ("����Ĳ���")
            oblog.showusererr
    End Select
    Set rs = server.CreateObject("Adodb.RecordSet")
    'response.Write(sql)
    rs.open sql, conn, 1, 1
    If rs.EOF And rs.bof Then
        strGuide = strGuide & " (����0ƪ" & tName & ")</h1>"
        response.write strGuide&"</div>"
    Else
        totalPut = rs.recordcount
        strGuide=strGuide & " (����" & totalPut & "ƪ" & tName&")</h1>"
        response.write strGuide
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > totalPut Then
            If (totalPut Mod MaxPerPage) = 0 Then
                CurrentPage = totalPut \ MaxPerPage
            Else
                CurrentPage = totalPut \ MaxPerPage + 1
            End If

        End If
        If CurrentPage = 1 Then
            showContent
            response.write oblog.showpage(strFileName, totalPut, MaxPerPage, True, True, "ƪ" & tName)
        Else
            If (CurrentPage - 1) * MaxPerPage < totalPut Then
                rs.Move (CurrentPage - 1) * MaxPerPage
                Dim bookmark
                bookmark = rs.bookmark
                showContent
                response.write oblog.showpage(strFileName, totalPut, MaxPerPage, True, True, "ƪ" & tName)
            Else
                CurrentPage = 1
                showContent
                response.write oblog.showpage(strFileName, totalPut, MaxPerPage, True, True, "ƪ" & tName)
            End If
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub
Sub showContent()
    Dim i, rssubject
    i = 0
%>
    <form name="myform" method="Post" action="user_blogmanage.asp?t=<%=t%>" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');" >
	
	<table>
		<tr class="tr01">
			<td class="tr02-1">ѡ��</td><td class="tr02-2">����</td><td class="tr02-3"><%=tName%>����</td><td class="tr02-4">����</td><td class="tr02-5">����ʱ��</td><td class="tr02-6">״̬</td><td class="tr02-7">����</td>
		</tr>
		

								

<%do while not rs.eof %>
<tr class="tr02">
    <td class="tr02-1"><input name='id' type='checkbox' id="id" value='<%=cstr(rs("logid"))%>' /></td>
    <td class="tr02-2"><%= oblog.requestID_selectName("oblog_logclass","classname","where classid="& Cint(rs("classid")))%><%'getsubname(rs("subjectid"),allsub)%></td>
    <td class="tr02-3">
<%if rs("logfile")<>"" then
    response.write ("<a href=" & rs("logfile") & " target=_blank>" & oblog.filt_html(rs("topic")) & " </a>")
Else
    response.write ("<font color=#D75372>" & oblog.filt_html(rs("topic")) &"</font>" )
end if%> </td>
    <td class="tr02-4"><%response.Write(oblog.filt_html(rs("author")))%></td>

    <td class="tr02-5" title="<%=rs("addtime")%>"><%if rs("addtime")<>"" then  response.write formatdatetime(rs("addtime"),2) else response.write "&nbsp;"%> </td>
    
    <td class="tr02-6"><%if rs("isdraft")=1 then response.Write("<font color=#D75372>�ݸ�</font>") else response.Write("����")%></td>
    <td class="tr02-7">
<%
response.write "<a href='user_blogmanage.asp?action=updatelog&id=" & rs("logid") & "&t=" & t & "'>����</a>&nbsp;"
response.write "<a href='user_post.asp?logid=" & rs("logid") & "&t=" & t & "'>�޸�</a>&nbsp;"
response.write "<a href='user_blogmanage.asp?action=del&id=" & rs("logid") & "&t=" & t & "' onClick='return confirm(""ȷ��Ҫɾ����"&tName&"��"");'>ɾ��</a>&nbsp;"
If oblog.logined_ulevel = 9 Then
	response.Write("<br />")
    If rs("isbest") = 1 Then
        response.write "<a href='user_blogmanage.asp?action=unisbest&id=" & rs("logid") & "&t=" & t & "'>ȡ��</a>&nbsp;"
    Else
        response.write "<a href='user_blogmanage.asp?action=isbest&id=" & rs("logid") & "&t=" & t & "'><font color='#038ad7'>�Ƽ�?</font></a>&nbsp;"
    End If
	If rs("passcheck") = 1 Then
        response.write "<a href='user_blogmanage.asp?action=Lock&id=" & rs("logid") & "&t=" & t & "'>ȡ�����</a>&nbsp;"
    Else
        response.write "<a href='user_blogmanage.asp?action=UnLock&id=" & rs("logid") & "&t=" & t & "'>ͨ�����</a>&nbsp;"
    End If
End If
%> </td>
</tr>

<%
    i = i + 1
    If i >= MaxPerPage Then Exit Do
    rs.movenext
Loop
%>
</table>

    <ul class="list_bot">
    <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" />
              ȫѡ
����:
    <input name="action" type="radio" value="del" checked onClick="document.myform.subject.disabled=true" />ɾ��&nbsp;&nbsp;&nbsp;&nbsp;
    <input name="action" type="radio" value="Move" onClick="document.myform.subject.disabled=false" />�ƶ���ר��
    <select name="subject" id="subject" disabled>
    <%=substr%>
    </select>&nbsp;&nbsp;&nbsp;
    <% if oblog.logined_ulevel=9 then %>
    <input name="action" type="radio" value="Lock" onClick="document.myform.UserLevel.disabled=true" />
    ��Ϊδ���&nbsp;&nbsp;
    <input name="action" type="radio" value="UnLock" onClick="document.myform.UserLevel.disabled=true" />
    ͨ�����&nbsp;&nbsp;&nbsp;<%end if%>
    <input type="submit" name="Submit" value=" ִ�� ">
    </ul>
    </form>
</div>
<%
End Sub

Sub delblog()
    If id = "" Then
        oblog.adderrstr ("��ָ��Ҫɾ����" & tName)
        oblog.showusererr
        Exit Sub
    End If
    If InStr(id, ",") > 0 Then
        Dim n, i
        id=FilterIDs(id)
        n = Split(id, ",")
        For i = 0 To UBound(n)
            deloneblog (n(i))
        Next
    Else
        deloneblog (id)
    End If
    oblog.showok "ɾ��" & tName & "�ɹ���", ""
End Sub

Sub deloneblog(logid)
    logid = CLng(logid)
    Dim uid, delname,rst,fso, sid,rst1
    Set rst = server.CreateObject("adodb.recordset")
	If Not IsObject(conn) Then link_database
    rst.open "select userid,logfile,subjectid,logtype from oblog_log where logid="&logid&wsql,conn,1,3
    If Not rst.EOF Then
        uid = rst(0)
        delname = Trim(rst(1))
        sid = rst(2)
        '����ͼƬ��¼
        If rst("logtype")=1 Then
        	Call DeletePhotos(logid)
   	 	End If
		'��ʵ������Ҫ���������ļ�����
		if true_domain=1 and delname<>"" then
			if instr(delname,"archives") then
				delname=right(delname,len(delname)-InstrRev(delname,"archives")+1)
			else
				delname=right(delname,len(delname)-InstrRev(delname,"/"))
			end if
			if oblog.logined_ulevel=9 then
				set rst1=oblog.execute("select user_dir,user_folder from oblog_user where userid="&clng(uid))
				if not rst1.eof then
					delname=rst1(0)&"/"&rst1(1)&"/"&delname
				end if
				set rst1=nothing
			else
				delname=oblog.logined_udir&"/"&oblog.logined_ufolder&"/"&delname
			end if
			'response.write(delname)
			'response.end
		end if
        If delname <> "" Then
            Set fso = server.CreateObject("Scripting.FileSystemObject")
            If fso.FileExists(server.MapPath(delname)) Then fso.DeleteFile server.MapPath(delname)
        End If
        rst.Delete
        rst.Close
        '---------------------------------
		Call Tags_UserDelete(logid)
        '--------------------------------------------
        '�˴���Ҫ��һ���޸ģ�������
        oblog.execute ("update oblog_user set log_count=log_count-1 where userid=" & uid)
        oblog.execute ("delete from oblog_comment where mainid=" & CLng(logid))
        oblog.execute ("update oblog_subject set subjectlognum=subjectlognum-1 where subjectid=" & CLng(sid))
        '--------------------------------------------
        Dim blog
        Set blog = New class_blog
        blog.userid = uid
        blog.Update_Subject uid
        blog.update_index 0
        blog.update_newblog (uid)
        Set blog = Nothing
        Set fso = Nothing
        Set rst = Nothing  
    Else
        rst.Close
        Set rst = Nothing
        Exit Sub
    End If
End Sub

Sub moveblog()
    If id = "" Then
        oblog.adderrstr ("��ָ��Ҫ�ƶ���" & tName)
        oblog.showusererr
        Exit Sub
    End If
    Dim subjectid
    subjectid = Trim(request("subject"))
    If subjectid = "" Then
        oblog.adderrstr ("��ָ��Ҫ�ƶ���Ŀ��ר��")
        oblog.showusererr
        Exit Sub
    Else
        subjectid = CLng(subjectid)
    End If
    If InStr(id, ",") > 0 Then
        id=FilterIDs(id)
        sql="Update [oblog_log] set subjectid="&subjectid&" where logid in (" & id & ")"&wsql
    Else
        sql="Update [oblog_log] set subjectid="&subjectid&" where logid=" & clng(id) &wsql
    End If
    oblog.execute sql
    Dim blog, rs1
    Set blog = New class_blog
    blog.userid = oblog.logined_uid
    blog.Update_Subject oblog.logined_uid
    'blog.update_index_subject 0, 0, 0, ""
    Set blog = Nothing
    Set rs = oblog.execute("select subjectid from oblog_subject where userid=" & oblog.logined_uid & " And Subjecttype=" & t)
    While Not rs.EOF
        Set rs1 = oblog.execute("select count(logid) from oblog_log where oblog_log.subjectid=" & rs(0))
        oblog.execute ("update oblog_subject set subjectlognum=" & rs1(0) & " where oblog_subject.subjectid=" & rs(0))
        rs.movenext
    Wend
    Set rs = Nothing
    Set rs1 = Nothing
    oblog.showok "����ר��ɹ�,��Ҫ���·�����ҳ���ſ�ʹר��ͳ��׼ȷ!", ""
End Sub

Sub updatelog()
    response.write ("<div id=""prompt""><ul>")
    id = CLng(id)
    Dim blog, p, rs, uid
    Set blog = New class_blog
    oblog.execute("update oblog_log set isdraft=0 where logid="&id&wsql)
    set rs=oblog.execute("select userid,subjectid from oblog_log where logid="&id&wsql)
    If Not rs.EOF Then
        blog.userid = rs(0)
        blog.progress_init
        p = 6
        blog.progress Int(1 / p * 100), "���ɾ�̬" & tName & "�ļ�"
        blog.update_log id, 1
        blog.progress Int(2 / p * 100), "����" & tName & "�ļ�"
        blog.update_calendar (id)
        blog.progress Int(3 / p * 100), "������ҳ"
        blog.update_index 0
        blog.progress Int(4 / p * 100), "����" & tName & "�����б�"
        blog.Update_Subject rs(0)
        blog.progress Int(5 / p * 100), "������" & tName & "�б�"
        blog.update_newblog (rs(0))
        blog.progress Int(6 / p * 100), "����" & tName & "���"
        Set rs = Nothing
		response.Write("<li><input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'></li></ul></div>")
    Else
        Set rs = Nothing
        oblog.adderrstr ("�޲���Ȩ��!")
        oblog.showusererr
    End If
End Sub

Sub isbest()
    id = CLng(id)
    oblog.execute("update oblog_log set isbest=1 where logid="&id&wsql)
    oblog.showok "��Ϊ�Ƽ��ɹ�,��ҳ����²���Ч!", ""
End Sub

Sub unisbest()
    id = CLng(id)
    oblog.execute("update oblog_log set isbest=0 where logid="&id&wsql)
    oblog.showok "ȡ���Ƽ��ɹ�,��ҳ����²���Ч!", ""
End Sub

Sub locklog()
    If id = "" Then
        oblog.adderrstr ("��ָ��Ҫȡ����˵�" & tName)
        oblog.showusererr
        Exit Sub
    End If
    If InStr(id, ",") > 0 Then
        Dim n, i
        id=FilterIDs(id)
        n = Split(id, ",")
        For i = 0 To UBound(n)
            lockonelog (n(i))
        Next
    Else
        lockonelog (id)
    End If
    oblog.showok "ȡ����˳ɹ���", ""
End Sub


Sub lockonelog(logid)
    logid = CLng(logid)
    Dim uid
    Set rs = server.CreateObject("adodb.recordset")
    rs.open "select userid,passcheck from oblog_log where logid="&logid&wsql,conn,1,3
    If Not rs.EOF Then
        rs(1) = 0
        uid = rs(0)
        rs.Update
        rs.Close
        Set rs = Nothing
        Dim blog
        Set blog = New class_blog
        blog.userid = uid
        blog.update_log logid, 0
        Set blog = Nothing
    Else
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
End Sub

Sub passlog()
    If id = "" Then
        oblog.adderrstr ("��ָ��Ҫ��˵�" & tName)
        oblog.showusererr
        Exit Sub
    End If
    If InStr(id, ",") > 0 Then
        Dim n, i
        id=FilterIDs(id)
        n = Split(id, ",")
        For i = 0 To UBound(n)
            passonelog (n(i))
        Next
    Else
        passonelog (id)
    End If
    oblog.showok "��˳ɹ���", ""
End Sub

Sub passonelog(logid)
    logid = CLng(logid)
    Dim uid
    Set rs = server.CreateObject("adodb.recordset")
    rs.open "select passcheck,userid from oblog_log where logid="&logid&wsql,conn,1,3
    If Not rs.EOF Then
        rs(0) = 1
        uid = rs(1)
        rs.Update
        rs.Close
        Set rs = Nothing
        Dim blog
        Set blog = New class_blog
        blog.userid = uid
        blog.update_log logid, 0
        blog.update_calendar (logid)
        Set blog = Nothing
    Else
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
End Sub

'ɾ����������ļ�
sub DeletePhotos(logid)
	dim rst,fs,fsize,uid,imgsrc,fid
	Set rst=oblog.execute("select file_path,file_size,userid,fileid from oblog_upfile where logid="&logid)
	if not rst.eof then
		set fs=createobject("scripting.filesystemobject")
		Do While Not rst.Eof
		fsize=rst(1)
		uid=rst(2)
		imgsrc=rst(0)
		fid=rst(3)
		if fs.FileExists(Server.MapPath(imgsrc)) then 
			fs.DeleteFile(server.mappath(imgsrc))
		end if
		if instr("jpg,bmp,gif,png,pcx",right(imgsrc,3))>0 then 'ɾ������ͼ
			imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
			imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
			if  fs.FileExists(Server.MapPath(imgsrc)) then
				fs.DeleteFile Server.MapPath(imgsrc)
			end if
		end if
		oblog.execute("delete from [oblog_upfile] where fileid="&fid)
		oblog.execute("update [oblog_user] set user_upfiles_size=user_upfiles_size-"&fsize&",user_upfiles_num=user_upfiles_num-1 where userid="&uid)
		rst.Movenext
		Loop
		set fs=nothing
		set rst=nothing
	end if	
end sub

Sub downlog()
%>
<style type="text/css">
<!--
    #list ul{ width: 96%;text-align: left;height: auto;padding:15px 3px 15px 3px;}
-->
</style>
 <h1 style='font-size:13px;'>��ѡ�����±��ݵ���ֹ���ڣ�</h1>
<div id="list">
     <!--<form name="oblogform" method="post" action="user_logzip.asp?t=<%=t%>" onSubmit="return submitdate();">-->
	 <form name="oblogform" method="post" action="user_logzip.asp?t=<%=t%>" onSubmit="return submitdate();">
    <ul>��ʼ���ڣ�<%oblog.type_dateselect dateadd("m",-1,date),1%></ul>
    <ul>�������ڣ�<%oblog.type_dateselect date(),2%></ul>
    <ul>���ݸ�ʽ��
          <input name="filetype" type="radio" value="txt" checked>
          TXT���ı�
           <input type="radio" name="filetype" value="htm">
          HTM��ҳ
          <input type="radio" name="filetype" value="xml">
          XML</ul>
    <ul><input type="submit" value=" �������� " /></ul>
    </form>
</div>
<%end sub

Function getsubname(sid, str)
	On Error Resume Next
    Dim tmp1, tmp2
	if sid<>"" then
		sid = CLng(sid)
	else
		sid=0
	end if
    If sid = 0 Then
        getsubname = "δ����"
        Exit Function
    End If
    tmp1=instr(str,sid&"!!??((")
    tmp2=len(sid&"!!??((")+tmp1
    If tmp1 > 0 Then
        getsubname = Mid(str, tmp2, InStr(tmp1, str, "##))==") - tmp2)
    Else
        getsubname = ""
    End If
End Function
%>
