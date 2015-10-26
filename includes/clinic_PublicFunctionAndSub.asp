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
public function filterhtml(str)
if trim(str)="" or isnull(str) then
 filterhtml=""
else
str=replace(str,">","&gt;")
str=replace(str,"<","&lt;")
str=replace(str,chr(32),"&nbsp;")
str=replace(str,chr(9),"&nbsp;")
str=replace(str,chr(34),"&quot;")
str=replace(str,chr(39),"&#39;")
str=replace(str,chr(13),"")
str=replace(str,chr(10)&chr(10),"</p><p>")
str=replace(str,chr(10),"<br>")
filterhtml=str
end if
end function
%>


<%
    Public Function JoinChar(strUrl)
        If strUrl = "" Then
            JoinChar = ""
            Exit Function
        End If
        If InStr(strUrl, "?") < Len(strUrl) Then
            If InStr(strUrl, "?") > 1 Then
                If InStr(strUrl, "&") < Len(strUrl) Then
                    JoinChar = strUrl & "&"
                Else
                    JoinChar = strUrl
                End If
            Else
                JoinChar = strUrl & "?"
            End If
        Else
            JoinChar = strUrl
        End If
    End Function

	Public Function Showpage_Front(sfilename, totalnumber, maxperpage, ShowTotal, ShowAllPages, strUnit)
        Dim n, i, strTemp, strUrl
        If totalnumber Mod maxperpage = 0 Then
            n = totalnumber \ maxperpage
        Else
            n = totalnumber \ maxperpage + 1
        End If
'        strTemp = "<div id='showpage' style='text-align:center;'>"
'strTemp = strTemp & "<Table><tr><td id='more'>"
		strTemp = "<tr><!--一个tr--> <td colspan='4' id='more' align='center' bgcolor='#ffffff'><!--一个td colspan='4' -->"
		
        If ShowTotal = True Then
            strTemp = strTemp & "共" & totalnumber & strUnit & "&nbsp;&nbsp;"
        End If
        strUrl = JoinChar(sfilename)
		'strUrl = sfilename---------------wl-------------------
        If CurrentPage < 2 Then
                strTemp = strTemp & "首页 上一页&nbsp;"
        Else
                strTemp = strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
                strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>上一页</a>&nbsp;"
        End If
    
        If n - CurrentPage < 1 Then
                strTemp = strTemp & "下一页 尾页"
        Else
                strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>下一页</a>&nbsp;"
                strTemp = strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
        End If
        strTemp = strTemp & "&nbsp;页次：" & CurrentPage & "/" & n & "</strong>页 "
        strTemp = strTemp & "&nbsp;" & maxperpage & "" & strUnit & "/页"
        If ShowAllPages = True Then
            strTemp = strTemp & "&nbsp;转到：<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"
            For i = 1 To n
                strTemp = strTemp & "<option value='" & i & "'"
                If CInt(CurrentPage) = CInt(i) Then strTemp = strTemp & " selected "
                strTemp = strTemp & ">" & i & "</option>"
            Next
            strTemp = strTemp & "</select>"
        End If
'strTemp = strTemp & "</td></tr></Table>"
'        strTemp = strTemp & "</div>"
		strTemp = strTemp & "</td><!--一个td colspan='4' --> </tr><!--一个tr-->"
        Showpage_Front = strTemp
    End Function

%>