<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<!--#include file="inc/inc_Tags.asp"-->
<%
Dim sSql,rst,iTagId,iUserId,sKeyword,i,sType,sAll
Dim sPage,sContent,sTitle,sForm,sErr
Dim oTagName,oNum,oLastUpdate,iReturn

If EN_TAGS=0 Then
	Call GoErr("ϵͳû������<font color=red>" & P_TAGS_DESC & "</font>����")
End If

sForm="<form name=""tagform"" method=""post"" action=""" & blogdir & P_TAGS_SYSURL & """ ><tr><td>" & VBCRLF
sForm = sForm & "&nbsp;&nbsp;" & P_TAGS_DESC & "�ؼ���&nbsp;&nbsp;<input type=""text"" name=""keyword"" size=20 value=""" & sKeyword & """>" & VBCRLF
sForm = sForm & "<input type=""submit"" value=""����"">" & VBCRLF
sForm = sForm & "</td></tr></form>" & VBCRLF

sType = LCase(Trim(Request.Querystring("t")))
iTagId = Trim(Request.Querystring("tagid"))
iUserId = Trim(Request.Querystring("userid"))
sKeyword= Trim(Request("keyword"))
sAll=Trim(Request.Querystring)
If sAll & sKeyword="" Then sType="hottags"

Call link_database()

Select Case sType
	Case "hottags"
		sTitle="�����ŵ�100��" & P_TAGS_DESC
		sContent=Tags_Hottags()
	Case "cloud"
		sTitle=P_TAGS_DESC & "��ͼ"
		sContent=Tags_SystemTags(1)
	Case "list"
		sTitle=P_TAGS_DESC & "�б�"
		sContent=Tags_SystemTags(0)
	Case "user"
		sTitle=P_TAGS_DESC & "�û�"
		sContent=GetUsersByTag(iTagId)
	Case Else
		If sKeyword="" Then
			If iTagId="" Then
				Call GoErr("����ָ��" & P_TAGS_DESC & " ID")
			Else
				If Not IsNumeric(iTagId) Then
					Call GoErr("�����" & P_TAGS_DESC & " ID")
				Else
					
					iReturn=CINT(Tags_TagName(iTagId,oTagName,oNum,oLastUpdate))					
					If iReturn=-1 Then																	
						oTagName="--"
					End If
					sTitle=P_TAGS_DESC & "&nbsp;:&nbsp;<font color=red>" & oTagName & "</font>"
					iTagId=CLNG(iTagId)						
				End If
			End If
			If iUserId="" Then
				sContent = Tags_TagBlogs("",iTagId)
			Else
				If Not IsNumeric(iUserId) Then
					Call GoErr("������û�ID")
				Else
					iUserId=CLNG(iUserId)
					iReturn=CINT(Tags_TagName(iTagId,oTagName,oNum,oLastUpdate))					
					If iReturn=-1 Then																	
						oTagName="--"
					End If
					sTitle="&nbsp;:&nbsp;<font color=red>" & GetUserInfo(iUserId) & "</font>," & P_TAGS_DESC & "&nbsp;:&nbsp;<font color=red>" & oTagName & "</font>"
					sContent = Tags_TagBlogs(iUserId,iTagId)
				End If
			End If
			If iUserId & iTagId ="" Then
				sTitle=P_TAGS_DESC & "��ͼ"
				sContent=Tags_SystemTags(1)
			End If
		Else
			sKeyword=ProtectSql(sKeyword)
			sTitle="����<font color=red>" & sKeyword & "</font>��" & P_TAGS_DESC
			sContent=Tags_SearchTag(sKeyword)
		End If
End Select

Function GoErr(byval sErrMsg)
	Response.Redirect "err.asp?message=" & sErrmsg
End Function

sPage = "<table width='100%' border='0'>" & VBCRLF
sPage = sPage & sForm  & VBCRLF
sPage = sPage & "<tr><td>��ǰλ�ã�<a href='index.asp'>��ҳ</a>��" & sTitle & VBCRLF
sPage = sPage & "</td><td align='right'><a href=""" & blogdir & P_TAGS_SYSURL & "?t=hottags"" >����" & P_TAGS_DESC &"</a>&nbsp;&nbsp;&nbsp;&nbsp; "
sPage = sPage & "<a href=""" & blogdir & P_TAGS_SYSURL & "?t=cloud"">" & P_TAGS_DESC &"��ͼ</a></td></tr></table><hr>" & VBCRLF
sPage = sPage & "<table width='100%' border='0'><tr><td>" & VBCRLF
sPage = sPage & sContent & VBCRLF
sPage = sPage & "</td></tr></table>" & VBCRLF

call sysshow()
show=replace(show,"$show_list$",sPage)
Response.Write show & oblog.site_bottom
If IsObject(conn) Then
	conn.close
	Set conn=Nothing
End If
%>
