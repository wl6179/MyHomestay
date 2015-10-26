<%
Option Explicit'强制变量声明；
Response.Buffer = true'缓冲数据，才向客户输出；
%>
<!--#include file="config.asp"-->
<%
Call SystemState
'修改系统配置请到config.asp文件
'此以下设置数据库及外部用户表整合参数
Const is_sqldata=1
'是否使用其他系统的数据库用户表,1为使用，0为不使用
Const is_ot_user=0
Dim connstr,conn,db,dchr,delchr
Dim ot_connstr,ot_conn,ot_usertable,ot_username,ot_password,ot_regurl,ot_lostpasswordurl,ot_modIfypass1,ot_modIfypass2
Sub link_database()
	Dim userDB
	If is_sqldata=0 Then
		'access数据库连接参数
		'第一次使用请修改本处数据库地址并相应修改data目录中数据库名称
		'如将oblog3.mdb修改为oblog3.asp
		'此处必须为以根目录开始,最前面必须为/号
		db =    "/oblog312/data/oblog31.mdb"
		connstr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)		
		dchr="#"
		delchr=" * "
	Else
		'sql数据库连接参数：数据库名、用户密码、用户名、连接名（本地用local，外地用IP）
		Dim sql_databasename,sql_password,sql_username,sql_localname
		sql_localname = "61.153.120.138"
		sql_databasename = "sq_myhomestay"
		sql_username = "sq_myhomestay"
		sql_password = "******"
		connstr = "Provider = Sqloledb; User ID = " & sql_username & "; Password = " & sql_password & "; Initial Catalog = " & sql_databasename & "; Data Source = " & sql_localname & ";"
		dchr="'"
		delchr=" "
	End If
	On Error Resume Next
	Set conn = Server.CreateObject("ADODB.Connection")	
	conn.open connstr

	rem ****若使用外部数据库表请自行修改下面的变量值****
	If is_ot_user=1 and instr(lcase(request.ServerVariables("HTTP_REFERER")),"admin_")=0 then 
		ot_connstr= "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath("/bbs/data/dvbbs71.mdb") 'access外部数据库连接字符串
		'ot_connstr = "Provider = Sqloledb; User ID = ***; Password = ***; Initial Catalog = ***; Data Source = ***;" 'sql外部数据库连接字符串
		Set ot_conn = Server.CreateObject("ADODB.Connection")
		ot_conn.open ot_connStr '外部数据库连接
		ot_usertable="dv_user" '用户表名
		ot_username="username" '用户名字段
		ot_password="userpassword" '密码字段
		ot_regurl="../bbs/reg.asp" '注册用户链接
		ot_modIfypass1="../bbs/modIfyadd.asp?t=1" '修改密码连接
		ot_modIfypass2="../bbs/modIfyadd.asp?t=1" '修改密码提示问题连接
		ot_lostpasswordurl="../bbs/lostpass.asp" '找回密码链接
	End If
	If err then
		err.clear
		Set conn = nothing
		response.write "数据库连接出错，请检查连接字串。"
		response.End
	End If
End Sub
%>
<!--#include file="Inc/Inc_Functions.asp"-->