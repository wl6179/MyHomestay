<!--#include file="conn.asp"-->
<html>
<body style="font-size:14px">
oblog 3.1 oblog_trackback��ṹ����
<BR/>
<BR/>
������ʱ�䣺2005��11��16�� 15:11
<BR/>
<BR/>
<input type="button" value="��ʼ����>>" onClick="window.location.href='update_trackback.asp?step=1'">
<BR/>
<BR/>
============================================
<BR/>
<BR/>
<%
Dim iStep,rs
iStep=Request("step")
'On error Resume Next
Dim strSql,strErrMsg,strNow

If iStep=1 Then
	Call Link_Database
	
	Call Execute("Alter Table oblog_trackback Add Ip varchar (255)  NULL")
	Call Execute("Alter Table oblog_trackback Add tb_url varchar (255)  NULL")
	Call Execute("Alter Table oblog_trackback Add blog_name varchar (255)  NULL")
	If strErrMsg<>"" Then
		Echo  "<font color=red>oblog_trackback �޸�ʧ��</font>:<BR/>" & strErrMsg  &"<BR/><BR/>" 
	Else
		Echo  "<font color=green>oblog_trackback �޸ĳɹ�</font><BR/><BR/>" 
	End If
	strErrMsg=""
	Response.Write "ִ�����!"
	Response.End
End If


%>

============================================
<BR/>
<BR/>
oBlog3.1(http://www.myhomestay.com.cn)
<BR/>
<BR/>
</body>
</html>
<%
'---------------------------------------------

Sub  Execute(sSql)
On Error Resume Next
	Call conn.Execute(sSql)
	If Err.Number<>0 Then
		strErrMsg = strErrMsg & "<font color=red>" & Err.Description & "</font><BR>"
	End If
	Err.Clear
	On Error Goto 0
End Sub

Sub Echo(sStr)
	Response.Write sStr
	Response.Flush
End Sub
%>