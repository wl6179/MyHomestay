
<%

	Dim userDB

		'sql���ݿ����Ӳ��������ݿ������û����롢�û�������������������local�������IP��
		Dim sql_databasename,sql_password,sql_username,sql_localname
		sql_localname = "202.75.219.208"
		sql_databasename = "Northwind"
		sql_username = "sq_myhomestay"
		sql_password = "noveli"
		connstr = "Provider = Sqloledb; User ID = " & sql_username & "; Password = " & sql_password & "; Initial Catalog = " & sql_databasename & "; Data Source = " & sql_localname & ";"
		dchr="'"
		delchr=" "

	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")	
	conn.open connstr


	If err then
		err.clear
		Set conn = nothing
		response.write "���ݿ����ӳ������������ִ���"
		response.End
	End If
	

%>
