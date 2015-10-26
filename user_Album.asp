<%@ LANGUAGE = VBScript CodePage = 936%>
<!--#include file="conn.asp"-->
<!--#include  file="inc/class_sys.asp"-->
<%
Dim oBlog
Dim CurrentPage,Maxperpage,strFileName,totalPut
Set oBlog = New class_sys
oBlog.start

'*********************************************************
'File: 			user_Album.asp
'Description:	用户相册列表 For oBlog3.1
'				暂用，可能整合到cmd.asp
'Author: 		王旭 
'Copyright:		http://www.myhomestay.com.cn
'LastUpdate:	20050921
'*********************************************************

If EN_PHOTO=0 then
	Response.Write "document.write('系统尚未启用相册功能。');"
	Response.End
End If

Dim userid,classid,fso,dsql,sPara,strAction
userid=Request("userid")
classid=Request("classid")
sPara=Request.QueryString
If userid<>"" Then
	If IsNumeric(userid) Then	
		userid=CLNG(userid)
	Else
		Response.Write "document.write('错误的用户ID');"
		Response.End
	End If		
Else
	Response.Write "document.write('错误的用户ID');"
	Response.End
End If

'strAction = "index.asp?userid=" & userid

If classid<>"" Then
	If IsNumeric(classid) Then	
		classid= CLNG(classid)
		If classid>0 Then
			dsql= " And userclassid=" & classid
			strAction = strAction & "&classid=" & classid
		End If
	End If
End If


Call update_albumIndex(userid,strAction)

Sub update_albumIndex(userid,strurl)
	Dim arrString,sContent
	Dim Sql,i,user_filepath,imgsrc
	Sql="select file_path,logid from oblog_upfile where isphoto=1 and isPower=0 and userid="&userid&dsql&"  order by fileid desc"
	sContent=showUserPhotolist(Sql,strurl)
	arrString=Split(sContent,vbcrlf)
	For i=0 To UBound(arrString)
		Response.Write("document.write('"&Replace(Replace(Replace(arrString(i),"\","\\"),"'","\'"),chr(13),"")&"');")
		Response.Write VBCRLF
		Response.Write("document.write('\n');")
		Response.Write VBCRLF
	Next
End Sub
	
Function showUserPhotolist(sql,strurl)
	Dim topn,msql,sReturn,rsPhoto
	MaxPerPage=oblog.Setup(79,0)
	strFileName=strurl
	If request("page")<>"" Then
    	currentPage=cint(request("page"))
	Else
		currentPage=1
	End If	
	If not IsObject(conn) Then link_database
	Set rsPhoto=Server.CreateObject("Adodb.RecordSet")
	rsPhoto.Open sql,Conn,1,1
  	If rsPhoto.eof and rsPhoto.bof then
		sReturn=sReturn & "共调用0个图片<br>"
	Else
    	totalPut=rsPhoto.recordcount
		If currentpage<1 then
       		currentpage=1
    	End If
    	If (currentpage-1)*MaxPerPage>totalput then
	   		If (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	Else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		End If

    	End If
	    If currentPage=1 then
        	sReturn = getPhotolist(rsPhoto)
        	sReturn=sReturn&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"个图片")
   	 	Else
   	     	If (currentPage-1)*MaxPerPage<totalPut then
         	   	rsPhoto.move  (currentPage-1)*MaxPerPage
         		Dim bookmark
           		bookmark=rsPhoto.bookmark
            	sReturn = getPhotolist(rsPhoto)
            	sReturn=sReturn&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"个图片")
        	Else
	        	currentPage=1
           		sReturn = getPhotolist(rsPhoto)
           		sReturn=sReturn&oblog.showpage(strFileName,totalput,MaxPerPage,false,true,"个图片")
	    	End If
		End If
	End If
	rsPhoto.Close
	Set rsPhoto=Nothing
	showUserPhotolist=sReturn
End Function

Function getPhotolist(rsPhoto)
	Dim i,bstr,n,fso,sReturn,strPlayerUrl
	Dim title,userurl,imgsrc
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	strPlayerUrl= blogdir & "PhotoPlayer.asp"
	sReturn="<table width='90%' border='0' style='font-size:12px'><tr><td>" & VBCRLF
	sReturn=sReturn&"<a href='index.asp'>相册首页</a>→共调用" & totalPut & " 个图片" & VBCRLF
	sReturn=sReturn&"</td><td align='right'><a href=""#"" onclick=""window.open('"&strPlayerUrl&"','_photo','height=500, width=480, top=100, left=400, toolbar=no, menubar=no, scrollbars=no, resizable=yes,status=no')"">启用自动播放</a></td></tr>" & VBCRLF
	sReturn=sReturn& GetUserClasses(strAction) & "</table><hr>"

	sReturn=sReturn&"<table width='100%'  align='center' cellpadding='0' cellspacing='1'>"& vbcrlf
	Do While not rsPhoto.eof
		sReturn=sReturn&"<tr height='22'>"& vbcrlf
		For n=1 to 4
			if rsPhoto.eof then
				sReturn=sReturn&"<td width='25%'></td>"& vbcrlf
			Else
				title="<BR/><a href=" & blogdir & "more.asp?id=" & rsPhoto(1) &" target=_blank style='font-size:12px'>阅读图片介绍</a><BR/><BR/>"
				userurl="<a href='index.asp' target='_blank'>"
				imgsrc=blogdir & rsPhoto(0)
				imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
				'imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
				if  not fso.FileExists(Server.MapPath(imgsrc)) then
					imgsrc=rsPhoto(0)
				End if
				sReturn=sReturn&"<td align='center'> <a href='"& blogdir & rsPhoto(0)&"' target='_blank' title='点击查看原图'><img src='"&imgsrc&"' height='100' width='130' border='0' /></a><br />"&title&"</td>"& vbcrlf
				i=i+1
				if not rsPhoto.eof then rsPhoto.movenext
			End if
		Next
		sReturn=sReturn&"</tr>"& vbcrlf
		if i>=MaxPerPage then exit do	
	loop		
	sReturn=sReturn&"</table>"	& VBCRLF
	Set fso=nothing
	getPhotolist=sReturn
End Function
	
'获取用户分类
Function GetUserClasses(strUrl)
	Dim rst,sReturn
	Set rst=conn.Execute("Select * From oblog_subject Where subjecttype=1 and userid=" & userid)
	If rst.Eof Then
		sReturn=""
	Else
		Do While Not rst.Eof
			sReturn="<option value="&rst("subjectid")&">" & rst("subjectname") & "</option>" & VBCRLF
			rst.Movenext
		Loop
		sReturn = "<option value="""">请选择分类</option><option value=0>所有分类</option>" & VBCRLF & sReturn
		'sReturn="<form method=get action=" & strUrl & "><tr><td clospan=2>选择相册分类：<select name=classid onpropertychange=""form.submit()"">" & VBCRLF & sReturn & "</select></tr></form>"
		sReturn="<form method=get><tr><td clospan=2>选择相册分类：<select name=classid onpropertychange=""form.submit()"">" & VBCRLF & sReturn & "</select></tr></form>"
	End If
	rst.Close
	Set rst=Nothing
	GetUserClasses = sReturn
End Function
%>