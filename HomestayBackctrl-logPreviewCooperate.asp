<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%

if not oblog.checkuserlogined_team() then
	'''response.Redirect("login.asp?fromurl="&replace(oblog.GetUrl,"&","$"))
	response.Redirect("index.asp?fromurl="&replace(oblog.GetUrl,"&","$"))
else
	if oblog.logined_ulevel=6 and oblog.logined_ulevel_team=0 then
		oblog.adderrstr("��δͨ��MyHomestay����Ա��ˣ����ܽ�����˹����̨")
		oblog.showerr
	end if
end if

if en_photo=0 then
	oblog.adderrstr("�˹����ѱ�ϵͳ�رգ�")
	oblog.showerr
	Set oblog = Nothing
end if
dim mainsql,strurl,rsmain,bstr1,isbest,show_list,classid,strPlayerUrl
strurl="photo.asp"
isbest=cint(request.QueryString("isbest"))
classid=request.QueryString("classid")

Call sysshow_3()

if isbest=1 then
	mainsql=mainsql&" and user_isbest=1"
	if strurl="photo.asp" then
		strurl=strurl&"?isbest=1"
	else
		strurl=strurl&"&isbest=1"
	end if
	bstr1="���Ƽ�ͼƬ"
end if

If IsNumeric(classid) Then
	classid=CLNG(classid)
	If classid>0 Then 
		mainsql= mainsql & " and sysClassid=" & classid & " "
		if strurl="photo.asp" then
			strurl=strurl&"?classid=" & classid
		else
			strurl=strurl&"?classid=" & classid
		end if
	End IF
End If

strPlayerUrl= Replace(strurl,"photo.asp","photoplayer.asp")
		
Call sub_showuserlist(mainsql,strurl)
show=replace(show,"$show_list$",show_list)
response.Write show&oblog.site_bottom

sub sub_showuserlist(sql,strurl)
	dim topn,msql
	MaxPerPage=oblog.setup(79,0)
	strFileName=strurl
	if request("page")<>"" then
    	currentPage=cint(request("page"))
	else
		currentPage=1
	end if
	msql="select top "&clng(oblog.setup(78,0))&" * from [oblog_logCooperateSubmit] where logid="& Cint(Request("logid")) &"  order by logid desc"
	if not IsObject(conn) then link_database
	Set rsmain=Server.CreateObject("Adodb.RecordSet")
	'response.Write(msql)
	rsmain.Open msql,Conn,1,1
  	if rsmain.eof and rsmain.bof then
		show_list=show_list & "������0��ͼƬ<br>"
	else
    	totalPut=rsmain.recordcount
		'show_list=show_list & "������" & totalPut & " ��ͼƬ<br>"
		if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if

    	end if
	    if currentPage=1 then
        	getlist()
        	'''show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"��ͼƬ")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rsmain.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rsmain.bookmark
            	getlist()
            	'''show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"��ͼƬ")
        	else
	        	currentPage=1
           		getlist()
           		'''show_list=show_list&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"��ͼƬ")
	    	end if
		end if
	end if
	rsmain.Close
	set rsmain=Nothing
end sub

sub getlist()
	dim i,bstr,n,fso
	dim title,userurl,imgsrc
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	
show_list= "��ǰλ�ã� ������ϵĴ�ŵ�Ч��Ԥ������ϸЧ�������ȴ�վ����˺�鿴<br />&nbsp;"	 
show_list= show_list & "<div style='margin-left:80px;width:470px;font-size:14px;'><H4>"
show_list= show_list & "<DIV class=show-topictxt style='text-align:center'><IMG src='images/Article.gif'>&nbsp;"& rsmain("topic") &"</H4></DIV>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "& rsmain("logtext")
show_list= show_list & "<HR color=#edeff0 size1>"
show_list= show_list & "&nbsp; Post &nbsp;by&nbsp;<A href='http://www.myhomestay.com.cn/' target=_blank>www.myhomestay.com.cn</A>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;������ "& rsmain("addtime") &"   "& oblog.WeekDateE(rsmain("addtime")) &"  &nbsp; "
show_list= show_list & "<HR color=#edeff0 size1></div>"

	
	
	set fso=nothing
end sub

'��ȡϵͳ����
Function GetSysClasses()
	Dim rst,sReturn
	Set rst=conn.Execute("Select * From oblog_logclass Where idtype=1")
	If rst.Eof Then
		sReturn=""
	Else
		Do While Not rst.Eof
			sReturn= sReturn & "<option value="&rst("id")&">" & rst("classname") & "</option>" & VBCRLF
			rst.Movenext
		Loop
		sReturn = "<option value=0>���з���</option>" & VBCRLF & sReturn
		sReturn="<form name=photoform method=get><tr><td clospan=2></tr></form>"
	End If
	rst.Close
	Set rst=Nothing
	GetSysClasses = sReturn
End Function
%>
