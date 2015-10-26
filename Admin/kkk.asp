<%=request.servervariables("http_accept_language")%>,
<%=request.servervariables("LOGON_USER")%> ,
<%=request.servervariables("CONTENT_TYPE")%>,
<%= Request.ServerVariables("http_host") %> ,<br /><br /><br />

......<br /><br /><br />
<font size="2" style="line-height:28px">
<%= Request.ServerVariables("Url") %>=
返回服务器地址<br />

<%= Request.ServerVariables("Path_Info") %>=
客户端提供的路径信息<br />

<%= Request.ServerVariables("Appl_Physical_Path") %>=
与应用程序元数据库路径相应的物理路径<br />

<%= Request.ServerVariables("Path_Translated") %>=
通过由虚拟至物理的映射后得到的路径<br />

<%= Request.ServerVariables("Script_Name") %>=
执行脚本的名称<br />

<%= Request.ServerVariables("Query_String") %>=
查询字符串內容<br />

<%= Request.ServerVariables("Http_Referer") %>=
请求的字符串內容<br />

<%= Request.ServerVariables("Server_Port") %>=
接受请求的服务器端口号<br />

<%= Request.ServerVariables("Remote_Addr") %>=
发出请求的远程主机的IP地址<br />

<%= Request.ServerVariables("Remote_Host") %>=
发出请求的远程主机名称<br />

<%= Request.ServerVariables("Local_Addr") %>=
返回接受请求的服务器地址<br />

<%= Request.ServerVariables("Http_Host") %>=
返回服务器地址<br />

<%= Request.ServerVariables("Server_Name") %>=
服务器的主机名、DNS地址或IP地址<br />

<%= Request.ServerVariables("Request_Method") %>=
提出请求的方法比如GET、HEAD、POST等等<br />

<%= Request.ServerVariables("Server_Port_Secure")%>=
如果接受请求的服务器端口为安全端口时，则为1，否则为0<br />

<%= Request.ServerVariables("Server_Protocol")%>=
服务器使用的协议的名称和版本<br />

<%= Request.ServerVariables("Server_Software")%>=
应答请求并运行网关的服务器软件的名称和版本<br />

<%= Request.ServerVariables("All_Http")%>=
客户端发送的所有HTTP标头，前缀HTTP_<br />

<%= Request.ServerVariables("All_Raw")%>=
客户端发送的所有HTTP标头,其结果和客户端发送时一样，没有前缀HTTP_<br />

<%= Request.ServerVariables("Appl_MD_Path")%>=
应用程序的元数据库路径<br />

<%= Request.ServerVariables("Content_Length")%>=
客户端发出內容的长度<br />

<%= Request.ServerVariables("Https")%>=
如果请求穿过安全通道（SSL），则返回ON如果请求来自非安全通道，则返回OFF<br />

<%= Request.ServerVariables("Instance_ID")%>=
IIS实例的ID号<br />

<%= Request.ServerVariables("Instance_Meta_Path")%>=
响应请求的IIS实例的元数据库路径<br />

<%= Request.ServerVariables("Http_Accept_Encoding")%>=
返回內容如：gzip,deflate<br />

<%= Request.ServerVariables("Http_Accept_Language")%>=
返回內容如：en-us<br />

<%= Request.ServerVariables("Http_Connection")%>=
返回內容：Keep-Alive<br />

<%= Request.ServerVariables("Http_Cookie")%>=
返回內容如：nVisiT%<br />

2DYum=125;ASPSESSIONIDCARTQTRA=FDOBFFABJGOECBBKHKGPFIJI;ASPSESSIONIDCAQQTSRB=LKJJPLABABILLPCOGJGAMKAM;ASPSESSIONIDACRRSSRA=DK

HHHFBBJOJCCONPPHLKGHPB<br />

<%= Request.ServerVariables("Http_User_Agent")%>=
返回內容：Mozilla/4.0(compatible;MSIE6.0;WindowsNT5.1;SV1)<br />

<%= Request.ServerVariables("Https_Keysize")%>=
安全套接字层连接关键字的位数，如128<br />

<%= Request.ServerVariables("Https_Secretkeysize")%>=
服务器验证私人关键字的位数如1024<br />

<%= Request.ServerVariables("Https_Server_Issuer")%>=
服务器证书的发行者字段<br />

<%= Request.ServerVariables("Https_Server_Subject")%>=
服务器证书的主题字段<br />

<%= Request.ServerVariables("Auth_Password")%>=
当使用基本验证模式时，客户在密码对话框中输入的密码<br />

<%= Request.ServerVariables("Auth_Type")%>=
是用户访问受保护的脚本时，服务器用於检验用户的验证方法<br />

<%= Request.ServerVariables("Auth_User")%>=
代证的用户名<br />

<%= Request.ServerVariables("Cert_Cookie")%>=
唯一的客户证书ID号<br />

<%= Request.ServerVariables("Cert_Flag")%>=
客户证书标誌，如有客户端证书，则bit0为0如果客户端证书验证无效，bit1被设置为1<br />

<%= Request.ServerVariables("Cert_Issuer")%>=
用户证书中的发行者字段<br />

<%= Request.ServerVariables("Cert_Keysize")%>=
安全套接字层连接关键字的位数，如128<br />

<%= Request.ServerVariables("Cert_Secretkeysize")%>=
服务器验证私人关键字的位数如1024<br />

<%= Request.ServerVariables("Cert_Serialnumber")%>=
客户证书的序列号字段<br />

<%= Request.ServerVariables("Cert_Server_Issuer")%>=
服务器证书的发行者字段<br />

<%= Request.ServerVariables("Cert_Server_Subject")%>=
服务器证书的主题字段<br />

<%= Request.ServerVariables("Cert_Subject")%>=
客户端证书的主题字段<br />

<%= Request.ServerVariables("Content_Type")%>=
客户发送的form內容或HTTPPUT的数据类型<br />
</font>
