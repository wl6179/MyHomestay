<%
Sub Select_General(Table,Option_ID,Option_name,Request_ID,Request_ID2,Where,OrderBy)
   '��5���������ֱ��ǣ�
    '������
    'ID����
    '��ʾ����
    '���ڶԱȵ��ֶ���[�ĸ��ֶ�ֵȥ�ȣ�]��
    'Request_ID2���ڶԱȵ�ֵ��
    '����'Where'������䣻
    '����'OrderBy'�����ֶΣ���
'���� �� --- Call Select_General("reg_Bed","classid","classname","classname","","Where Ward_ID="&wardid,"Order By RootID,OrderID")
dim scID_Rs
set scID_Rs=conn.execute("select * from "&Table&" "&Where&" "&OrderBy)  
'wl   �޸�Ŀ�ģ� Order by truename Desc
 If scID_Rs.eof Then
  'response.write "��δ��ҽ����¼"
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

	'���� �� --- Call Select_General_BothID("clinic_Week","clinic_Time","clinicWeekID","clinicTimeID","classname","classname",Request("clinicWeekID"),Request("clinicTimeID"),"/*Where01*/","/*Where02*/","/*OrderBy01*/","/*OrderBy02*/")
dim scID_Rs01,scID_Rs02
set scID_Rs01=conn.execute("select * from "&Table01&" "&Where01&" "&OrderBy01)  
set scID_Rs02=conn.execute("select * from "&Table02&" "&Where02&" "&OrderBy02)  

 If scID_Rs01.eof Then
  'response.write "��δ���ڵļ�¼"
'  	   <option value="clinicWeekID=1&amp;clinicTimeID=">��һ</option>
'	<option value="clinicWeekID=1&amp;clinicTimeID=1">--��һ����</option>
'	<option value="clinicWeekID=1&amp;clinicTimeID=2">--��һ����</option>
'	<option value="clinicWeekID=1&amp;clinicTimeID=4">--��һҹ��</option>
 Else
  while not scID_Rs01.eof 
   %>
	   <option value="clinicWeekID=<%=scID_Rs01(Option_ID01)%>&amp;clinicTimeID=" <%If scID_Rs01(Option_ID01) = Request_ID_01 Then response.write " selected"%>><%=scID_Rs01(Option_name01)%></option>
	   
   			 
			 <%
			 If scID_Rs02.eof Then
			  'response.write "��δ���ڵļ�¼"
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
  
  
 

<%'��ȡ [ĳID] ����Ӧ�� [ĳ�ֶ�ֵ]~
FUNCTION requestID_selectName(Table,Field_name,Where)
   '��3���������ֱ��ǣ�
    '������
    'ID����
    '��ʾ����
    '����'Where'������䣻
'���� �� --- If Bed_ID<>0 Then Bed_Name = requestID_selectName("reg_Bed","classname","Where classid="&Bed_ID)
dim scID_Rs_Temp
set scID_Rs_Temp=conn.execute("select "&Field_name&" from "&Table&" "&Where)  
 If scID_Rs_Temp.eof Then
  requestID_selectName = "û�д˼�¼"
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
		strTemp = "<tr><!--һ��tr--> <td colspan='4' id='more' align='center' bgcolor='#ffffff'><!--һ��td colspan='4' -->"
		
        If ShowTotal = True Then
            strTemp = strTemp & "��" & totalnumber & strUnit & "&nbsp;&nbsp;"
        End If
        strUrl = JoinChar(sfilename)
		'strUrl = sfilename---------------wl-------------------
        If CurrentPage < 2 Then
                strTemp = strTemp & "��ҳ ��һҳ&nbsp;"
        Else
                strTemp = strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
                strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>��һҳ</a>&nbsp;"
        End If
    
        If n - CurrentPage < 1 Then
                strTemp = strTemp & "��һҳ βҳ"
        Else
                strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>��һҳ</a>&nbsp;"
                strTemp = strTemp & "<a href='" & strUrl & "page=" & n & "'>βҳ</a>"
        End If
        strTemp = strTemp & "&nbsp;ҳ�Σ�" & CurrentPage & "/" & n & "</strong>ҳ "
        strTemp = strTemp & "&nbsp;" & maxperpage & "" & strUnit & "/ҳ"
        If ShowAllPages = True Then
            strTemp = strTemp & "&nbsp;ת����<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"
            For i = 1 To n
                strTemp = strTemp & "<option value='" & i & "'"
                If CInt(CurrentPage) = CInt(i) Then strTemp = strTemp & " selected "
                strTemp = strTemp & ">" & i & "</option>"
            Next
            strTemp = strTemp & "</select>"
        End If
'strTemp = strTemp & "</td></tr></Table>"
'        strTemp = strTemp & "</div>"
		strTemp = strTemp & "</td><!--һ��td colspan='4' --> </tr><!--һ��tr-->"
        Showpage_Front = strTemp
    End Function

%>