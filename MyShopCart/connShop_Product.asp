
<%

	Dim userDB

		'sql数据库连接参数：数据库名、用户密码、用户名、连接名（本地用local，外地用IP）
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
		response.write "数据库连接出错，请检查连接字串。"
		response.End
	End If
	

%>
