<% 
'--------���岿�� ------------------ 
'�������վ��ע������DW���ɵ�ConnectionsĿ¼�µ����ݿ������ļ������ͷ�����û�ֱ���������������
'��Ҫע����ǣ�����ӵ����ݿ������ļ��У������ں�̨��������µ�����ʱ��������SQL���ϵͳ������ΪSQL��������ʾ����

'Ҫ�����SQLע��ʽ��������Ҫ�ֹ��޸ģ� 
'ֻҪ������ĳ��򱣴�ΪSQLinjection.asp,Ȼ������Ҫ��ע���ҳ��ͷ������ 
'< ! - - # Include File="/includes/SQLinjection.asp"- - > 
'�Ϳ�������ҳ���ע��.

dim sql_injdata 
SQL_injdata = "'|and|exec|insert|select|delete
|update|count|*|%|chr|mid|master|truncate|char
|declare|1=1|1=2|;" 
SQL_inj = split(SQL_Injdata,"|" 

'--------POST����------------------ 
If Request.QueryString<>"" Then 
For Each SQL_Get In Request.QueryString 
For SQL_Data=0 To Ubound(SQL_inj) 
if instr(Request.QueryString(SQL_Get),
Sql_Inj(Sql_DATA))>0 Then 
Response.Write "<Script Language=JavaScript>
alert('ϵͳ��ʾ��!\n\n�벻Ҫ�ڲ����а����Ƿ��ַ�����ע��!\n\n');window.location="&"'"&"index.htm"&"'"&";</Script>" 
Response.end 
end if 
next 
Next 
End If 

'--------GET����------------------- 
If Request.Form<>"" Then 
For Each Sql_Post In Request.Form 
For SQL_Data=0 To Ubound(SQL_inj) 
if instr(Request.Form(Sql_Post),Sql_Inj(Sql_DATA))>0 Then 
Response.Write "<Script Language=JavaScript>
alert('ϵͳ��ʾ��!\n\n�벻Ҫ�ڲ����а����Ƿ��ַ�����ע��!\n\n');window.location="&"'"&"index.htm"&"'"&";</Script>" 
Response.end 
end if 
next 
next 
end if 
%> 

