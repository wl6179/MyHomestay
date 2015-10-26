<%
Sub Select_General(Table,Option_ID,Option_name,Request_ID,Request_ID2,Where,OrderBy)
   '有5个参数，分别是：
    '表名；
    'ID名；
    '显示名；
    '用于对比的字段名[哪个字段值去比！]；
    'Request_ID2用于对比的值；
    '加上'Where'条件语句；
    '加上'OrderBy'排序字段；；
'例子 — --- Call Select_General("reg_Bed","classid","classname","classname","","Where Ward_ID="&wardid,"Order By RootID,OrderID")
dim scID_Rs
set scID_Rs=conn.execute("select * from "&Table&" "&Where&" "&OrderBy)  
'wl   修改目的： Order by truename Desc
 If scID_Rs.eof Then
  'response.write "尚未有医生记录"
 Else
  while not scID_Rs.eof 
   %>
   <option value=<%=scID_Rs(Option_ID)%>
  <%If scID_Rs(Request_ID) = Request_ID2 Then response.write " selected"%>>
 <%=scID_Rs(Option_name)%> <%'=scID_Rs(Option_ID)%></option>
  <%scID_Rs.movenext
  wend%>
  
 <%
 End If
End Sub
  %>



<%
Sub Select_General_BothID(Table01,Table02,Option_ID01,Option_ID02,Option_name01,Option_name02,Request_ID_01,Request_ID_02,Where01,Where02,OrderBy01,OrderBy02)

	'例子 — --- Call Select_General_BothID("clinic_Week","clinic_Time","clinicWeekID","clinicTimeID","classname","classname",Request("clinicWeekID"),Request("clinicTimeID"),"/*Where01*/","/*Where02*/","/*OrderBy01*/","/*OrderBy02*/")
dim scID_Rs01,scID_Rs02
set scID_Rs01=conn.execute("select * from "&Table01&" "&Where01&" "&OrderBy01)  
set scID_Rs02=conn.execute("select * from "&Table02&" "&Where02&" "&OrderBy02)  

 If scID_Rs01.eof Then
  'response.write "尚未星期的记录"
'  	   <option value="clinicWeekID=1&amp;clinicTimeID=">周一</option>
'	<option value="clinicWeekID=1&amp;clinicTimeID=1">--周一上午</option>
'	<option value="clinicWeekID=1&amp;clinicTimeID=2">--周一下午</option>
'	<option value="clinicWeekID=1&amp;clinicTimeID=4">--周一夜间</option>
 Else
  while not scID_Rs01.eof 
   %>
	   <option value="clinicWeekID=<%=scID_Rs01(Option_ID01)%>&amp;clinicTimeID=" <%If scID_Rs01(Option_ID01) = Request_ID_01 Then response.write " selected"%>><%=scID_Rs01(Option_name01)%></option>
	   
   			 
			 <%
			 If scID_Rs02.eof Then
			  'response.write "尚未星期的记录"
			 Else
			 
			  while not scID_Rs02.eof 
			   %>
				   <%IF scID_Rs02(Option_ID02)=3 THEN%>
				   <%'''''''''''''''''''''''''''''''''%>
				   <%ELSE%>
					   <option value="clinicWeekID=<%=scID_Rs01(Option_ID01)%>&amp;clinicTimeID=<%=scID_Rs02(Option_ID02)%>"
					   
					  <%If scID_Rs01(Option_ID01) = Request_ID_01 And scID_Rs02(Option_ID02) = Request_ID_02 Then response.write " selected"%>>
					  --<%=scID_Rs01(Option_name01)%><%=scID_Rs02(Option_name02)%>
					  
					   </option>
				   <%END IF%>
			  <%scID_Rs02.movenext
			  wend
			  scID_Rs02.moveFirst
			 
			 End If%>
   

	  
  <%scID_Rs01.movenext
  wend%>
  
 <%
 End If
End Sub
  %>
  
  
 

<%'获取 [某ID] 所对应的 [某字段值]~
FUNCTION requestID_selectName(Table,Field_name,Where)
   '有3个参数，分别是：
    '表名；
    'ID名；
    '显示名；
    '加上'Where'条件语句；
'例子 — --- If Bed_ID<>0 Then Bed_Name = requestID_selectName("reg_Bed","classname","Where classid="&Bed_ID)
dim scID_Rs_Temp
set scID_Rs_Temp=conn.execute("select "&Field_name&" from "&Table&" "&Where)  
 If scID_Rs_Temp.eof Then
  requestID_selectName = "没有此记录"
 Else
  requestID_selectName = scID_Rs_Temp(Field_name)
 
 End If
End FUNCTION
%>



<%
'放在根目录下 /inc/class_sys.asp 里边了！！！wl---------------------------

'public function filterhtml(str)
'if trim(str)="" or isnull(str) then
' filterhtml=""
'else
'str=replace(str,">","&gt;")
'str=replace(str,"<","&lt;")
'str=replace(str,chr(32),"&nbsp;")
'str=replace(str,chr(9),"&nbsp;")
'str=replace(str,chr(34),"&quot;")
'str=replace(str,chr(39),"&#39;")
'str=replace(str,chr(13),"")
'str=replace(str,chr(10)&chr(10),"</p><p>")
'str=replace(str,chr(10),"<br>")
'filterhtml=str
'end if
'end function
%>
