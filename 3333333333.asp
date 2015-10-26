<%
Dim server_v1, server_v2,ChkPost
		ChkPost = False
		if true_domain=1 then
			ChkPost=true
			'exit function
        end if
        server_v1 = CStr(request.ServerVariables("HTTP_REFERER"))
        server_v2 = CStr(request.ServerVariables("SERVER_NAME"))
        If Mid(server_v1, 8, Len(server_v2)) = server_v2 Then ChkPost = True'½ØÈ¡×Ö·ûÊýLen(server_v2)£»

response.Write request.ServerVariables("HTTP_REFERER")&";"
response.Write request.ServerVariables("SERVER_NAME")&";"

response.Write ChkPost
%>