<%
Dim conn_OldSite
set conn_OldSite=server.createobject("adodb.connection")
conn_OldSite.open "Provider = Sqloledb; User ID = 'sq_myhomestay'; Password = 'noveli'; Initial Catalog = sq_myhomestay; Data Source = '61.153.120.138';"
''conn_OldSite.open "driver={SQL server};server=61.153.120.138;uid=sq_myhomestay;pwd=noveli;database=sq_myhomestay"
%>