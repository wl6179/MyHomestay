<% 
'--------定义部份 ------------------ 
'如果想整站防注，就在DW生成的Connections目录下的数据库连接文件中添加头部调用或直接添加下面程序代码
'需要注意的是，在添加到数据库连接文件中，可能在后台表单添加文章等内容时，如遇到SQL语句系统会误认为SQL攻击而提示出错。

'要想防范SQL注入式攻击就需要手工修改， 
'只要将下面的程序保存为SQLinjection.asp,然后在需要防注入的页面头部调用 
'< ! - - # Include File="/includes/SQLinjection.asp"- - > 
'就可以做到页面防注入.

dim sql_injdata 
SQL_injdata = "'|and|exec|insert|select|delete
|update|count|*|%|chr|mid|master|truncate|char
|declare|1=1|1=2|;" 
SQL_inj = split(SQL_Injdata,"|" 

'--------POST部份------------------ 
If Request.QueryString<>"" Then 
For Each SQL_Get In Request.QueryString 
For SQL_Data=0 To Ubound(SQL_inj) 
if instr(Request.QueryString(SQL_Get),
Sql_Inj(Sql_DATA))>0 Then 
Response.Write "<Script Language=JavaScript>
alert('系统提示你!\n\n请不要在参数中包含非法字符尝试注入!\n\n');window.location="&"'"&"index.htm"&"'"&";</Script>" 
Response.end 
end if 
next 
Next 
End If 

'--------GET部份------------------- 
If Request.Form<>"" Then 
For Each Sql_Post In Request.Form 
For SQL_Data=0 To Ubound(SQL_inj) 
if instr(Request.Form(Sql_Post),Sql_Inj(Sql_DATA))>0 Then 
Response.Write "<Script Language=JavaScript>
alert('系统提示你!\n\n请不要在参数中包含非法字符尝试注入!\n\n');window.location="&"'"&"index.htm"&"'"&";</Script>" 
Response.end 
end if 
next 
next 
end if 
%> 

