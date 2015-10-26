<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%
Dim oFSO
Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
If Application(oblog.cache_name&"index_update")=false and oFSO.FileExists(Server.mappath("index.html")) and DateDiff("s",Application(oblog.cache_name&"index_updatetime"),Now())<oblog.setup(17,0) Then
	Set oFSO = Nothing
	Response.Redirect("index.html")
Else
	Set oFSO = Nothing
	Dim rstmp
	Set rstmp=oblog.execute("select skinmain from oblog_sysskin where isdefault=1")
	show=show&rstmp(0)
	set rstmp=nothing
	Call indexshow()
	show=show&oblog.site_bottom		
	Call oblog.BuildFile(Server.mappath("index.html"),show)	
	Application.Lock
		Application(oblog.cache_name&"index_update")=false
		Application(oblog.cache_name&"index_updatetime")=now()
		Application(oblog.cache_name&"list_update")=true
		Application(oblog.cache_name&"class_update")=false
	Application.unLock
	If Request("re")<>"0" Then
		Response.Redirect("index.html")
	End If
End If
%>