<%
Option Explicit'ǿ�Ʊ���������
Response.Buffer = true'�������ݣ�����ͻ������
%>
<!--#include file="config.asp"-->
<%
Call SystemState
'�޸�ϵͳ�����뵽config.asp�ļ�
'�������������ݿ⼰�ⲿ�û������ϲ���
Const is_sqldata=1
'�Ƿ�ʹ������ϵͳ�����ݿ��û���,1Ϊʹ�ã�0Ϊ��ʹ��
Const is_ot_user=0
Dim connstr,conn,db,dchr,delchr
Dim ot_connstr,ot_conn,ot_usertable,ot_username,ot_password,ot_regurl,ot_lostpasswordurl,ot_modIfypass1,ot_modIfypass2
Sub link_database()
	Dim userDB
	If is_sqldata=0 Then
		'access���ݿ����Ӳ���
		'��һ��ʹ�����޸ı������ݿ��ַ����Ӧ�޸�dataĿ¼�����ݿ�����
		'�罫oblog3.mdb�޸�Ϊoblog3.asp
		'�˴�����Ϊ�Ը�Ŀ¼��ʼ,��ǰ�����Ϊ/��
		db =    "/oblog312/data/oblog31.mdb"
		connstr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)		
		dchr="#"
		delchr=" * "
	Else
		'sql���ݿ����Ӳ��������ݿ������û����롢�û�������������������local�������IP��
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

	rem ****��ʹ���ⲿ���ݿ���������޸�����ı���ֵ****
	If is_ot_user=1 and instr(lcase(request.ServerVariables("HTTP_REFERER")),"admin_")=0 then 
		ot_connstr= "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath("/bbs/data/dvbbs71.mdb") 'access�ⲿ���ݿ������ַ���
		'ot_connstr = "Provider = Sqloledb; User ID = ***; Password = ***; Initial Catalog = ***; Data Source = ***;" 'sql�ⲿ���ݿ������ַ���
		Set ot_conn = Server.CreateObject("ADODB.Connection")
		ot_conn.open ot_connStr '�ⲿ���ݿ�����
		ot_usertable="dv_user" '�û�����
		ot_username="username" '�û����ֶ�
		ot_password="userpassword" '�����ֶ�
		ot_regurl="../bbs/reg.asp" 'ע���û�����
		ot_modIfypass1="../bbs/modIfyadd.asp?t=1" '�޸���������
		ot_modIfypass2="../bbs/modIfyadd.asp?t=1" '�޸�������ʾ��������
		ot_lostpasswordurl="../bbs/lostpass.asp" '�һ���������
	End If
	If err then
		err.clear
		Set conn = nothing
		response.write "���ݿ����ӳ������������ִ���"
		response.End
	End If
End Sub
%>
<!--#include file="Inc/Inc_Functions.asp"-->