<!--#include file="conn.asp" -->
<!--#include file="Inc/Class_TrackBack.asp" -->
<%
'*********************************************************
'File: 			tb.asp
'Description:	Receive Ping For oBlog3.1
'Author: 		ÍõÐñ 
'LastUpdate:	20050926
'
'Copyright:		http://www.myhomestay.com.cn
'*********************************************************
'ON Error Resume Next
Dim objTrackback
Set objTrackback = New Class_TrackBack
objTrackback.LOGID=Request("id")
objTrackback.IP=GetIP
objTrackback.URL=Trim(Request("url"))
'objTrackback.TBUSER=Trim(Request.QueryString("tbuser"))
objTrackback.TITLE=Trim(Request("title"))
objTrackback.BLOG_NAME=Trim(Request("blog_name"))
objTrackback.EXCERPT=Trim(Request("excerpt"))
Call Link_Database
Call objTrackback.Receive()
Set objTrackback=Nothing
conn.Close
Set conn=Nothing
%>