<%
'��ȡ���ڵ���(ʡ/��)����ѧϰ�������ࡪ�����ݽṹ������ͥ��Ų�ѯ��
Function GetSysClassesSearchHome_start()
	Dim strType_city
	Dim str_usertype
	Dim rst,sReturn
	
	Dim Request_province,Request_city
	Request_province= Request("province")
	Request_city= Request("city")
	
	Dim Request_usertype
	Request_usertype= Cint(Request("usertype"))
	
	Dim Request_houseframe
	Request_houseframe= Trim(Request("houseframe"))
	
	strType_city =strType_city & "<Div id='line01' style='margin:0px;padding-top:3px;'>&nbsp;���ڵ���&nbsp;" &oblog.type_citySelectSubmit(Request_province,Request_city )'����ѡ�������б�
	
	
	str_usertype="<select name=usertype id=usertype onchange='submit();'>"
	str_usertype=str_usertype&oblog.show_class("user",Request_usertype,0)
    str_usertype=str_usertype&"</select>"	
	str_usertype = "&nbsp;�����ֲ�ѯ&nbsp;" &str_usertype&"</Div>"'����ѡ�������б�
	
'	Set rst=conn.Execute("Select * From oblog_logclass Where idtype=1")
'	If rst.Eof Then
'		sReturn=""
'	Else
'		Do While Not rst.Eof
'			sReturn= sReturn & "<option value="&rst("id")&">" & rst("classname") & "</option>" & VBCRLF
'			rst.Movenext
'		Loop
'		sReturn = "<option value=0>���з���</option>" & VBCRLF & sReturn
		sReturn="<form name=selectsubmitform action='HomestayFamilyChina.asp' method=get><tr><td clospan=2>" & strType_city & str_usertype & "<Div id='line02' style='margin:0px;padding-top:6px;'>&nbsp;���ݽṹ&nbsp;<select name='houseframe' id='houseframe1' onchange='submit();' title='��ѡ�����ķ��ݽṹ'>"&_
                            "<option value='' selected>��</option>"&_
							"<option value='1��'"
							If Instr(Request_houseframe,"1��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">1��</option>"&_
                            "<option value='2��'"
							If Instr(Request_houseframe,"2��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">2��</option>"&_
                            "<option value='3��'"
							If Instr(Request_houseframe,"3��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">3��</option>"&_
                            "<option value='4��'"
							If Instr(Request_houseframe,"4��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">4��</option>"&_
                            "<option value='5��'"
							If Instr(Request_houseframe,"5��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">5��</option>"&_
                            "<option value='6��'"
							If Instr(Request_houseframe,"6��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">6��</option>"&_
                          "</select>&nbsp;<select name='houseframe' id='houseframe2' onchange='submit();' title='��ѡ�����ķ��ݽṹ'>"&_
						  "<option value='' selected>��</option>"&_
                            "<option value='1��'"
							If Instr(Request_houseframe,"1��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">1��</option>"&_
                            "<option value='2��'"
							If Instr(Request_houseframe,"2��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">2��</option>"&_
                            "<option value='3��'"
							If Instr(Request_houseframe,"3��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">3��</option>"&_
                            "<option value='4��'"
							If Instr(Request_houseframe,"4��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">4��</option>"&_
                            "<option value='5��'"
							If Instr(Request_houseframe,"5��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">5��</option>"&_
                            "<option value='6��'"
							If Instr(Request_houseframe,"6��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">6��</option>"&_
                          "</select>&nbsp;<select name='houseframe' id='houseframe3' onchange='submit();' title='��ѡ�����ķ��ݽṹ'>"&_
						  "<option value='' selected>��</option>"&_
                            "<option value='1��'"
							If Instr(Request_houseframe,"1��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">1��</option>"&_
                            "<option value='2��'"
							If Instr(Request_houseframe,"2��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">2��</option>"&_
                            "<option value='3��'"
							If Instr(Request_houseframe,"3��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">3��</option>"&_
                            "<option value='4��'"
							If Instr(Request_houseframe,"4��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">4��</option>"&_
                            "<option value='5��'"
							If Instr(Request_houseframe,"5��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">5��</option>"&_
                            "<option value='6��'"
							If Instr(Request_houseframe,"6��")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">6��</option>"&_
                          "</select></Div> </tr></form>"
'	End If
'	rst.Close
'	Set rst=Nothing
	GetSysClassesSearchHome_start = sReturn
End Function

GetSysClassesSearchHome_start()
%>
123