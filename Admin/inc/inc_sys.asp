<!--#include file="../../conn.asp"-->
<!--#include file="../../inc/class_sys.asp"-->
<!--#include file="../../inc/md5.asp"-->
<%
dim oblog
set oblog=new class_sys
oblog.start
chk_sysadmin
dim admin_name
sub chk_sysadmin()
	dim admin_password,sql,rs	
	admin_name=oblog.filt_badstr(session("adminname"))
	admin_password=oblog.filt_badstr(session("adminpassword"))
	if admin_name="" then
		response.redirect "administratorLogin.asp"
		exit sub
	end if
	sql="select id from oblog_admin where username='" & admin_name & "' and password='"&admin_password&"'"
	If Not IsObject(conn) Then link_database
	set rs=conn.execute(sql)
	if rs.bof and rs.eof then
		set rs=nothing
		response.redirect "administratorLogin.asp"
		exit sub
	end if
	rs.close
	set rs=nothing		
end sub
Function CheckSafePath(byval strMode)
	Dim strPathFrom,strPathSelf,arrFrom,arrSelf,i
	CheckSafePath=False
	If oBlog.ChkPost=False Then Exit Function
	strPathFrom = Replace(LCase(CStr(request.ServerVariables("HTTP_REFERER"))),"http://","")
    strPathSelf = Replace(LCase(CStr(request.ServerVariables("URL"))),"http://","")
    If strPathFrom="" Then Exit Function
    If strPathSelf="" Then Exit Function
    arrFrom=Split(strPathFrom,"/")
    arrSelf=Split(strPathSelf,"/")
    For i=0 To UBound(arrFrom)
    	'Response.Write "arrFrom("&i&")="& arrFrom(i) & "<BR/>" 
	Next 
	For i=0 To UBound(arrSelf)
    	'Response.Write "arrSelf("&i&")="& arrSelf(i) & "<BR/>"
	Next 
    Select Case strMode
    	Case "0"
			 For i = 1 To (UBound(arrSelf)-1)
             If arrFrom(i)=arrSelf(i) And Left(arrFrom(UBound(arrfrom)),Len(arrSelf(UBound(arrself))))=arrSelf(UBound(arrself))Then CheckSafePath=True  
			 Next 
			 'If arrFrom(1)=arrSelf(1) And Left(arrFrom(2),Len(arrSelf(2)))=arrSelf(2)Then CheckSafePath=True
			 'Response.Write checkSafePath
			 End Select
End Function
%>
<SCRIPT LANGUAGE="JavaScript">
function CheckSel(Voption,Value)
{
	var obj = document.getElementById(Voption);
	for (i=0;i<obj.length;i++){
		if (obj.options[i].value==Value){
		obj.options[i].selected=true;
		break;
		}
	}
}
function chang_size(num,objname)
{
	var obj=document.getElementById(objname)
	if (parseInt(obj.rows)+num>=3) {
		obj.rows = parseInt(obj.rows) + num;	
	}
	if (num>0)
	{
		obj.width="90%";
	}
}
</script>
<style type="text/css">
#showpage{
	CLEAR: both;
	text-align: center;
	width: 100%;
}
</style>

