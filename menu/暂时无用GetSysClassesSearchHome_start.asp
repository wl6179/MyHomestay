<%
'获取所在地区(省/市)——学习语言种类——房屋结构——家庭编号查询！
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
	
	strType_city =strType_city & "<Div id='line01' style='margin:0px;padding-top:3px;'>&nbsp;所在地区&nbsp;" &oblog.type_citySelectSubmit(Request_province,Request_city )'城市选择下拉列表；
	
	
	str_usertype="<select name=usertype id=usertype onchange='submit();'>"
	str_usertype=str_usertype&oblog.show_class("user",Request_usertype,0)
    str_usertype=str_usertype&"</select>"	
	str_usertype = "&nbsp;按语种查询&nbsp;" &str_usertype&"</Div>"'语种选择下拉列表；
	
'	Set rst=conn.Execute("Select * From oblog_logclass Where idtype=1")
'	If rst.Eof Then
'		sReturn=""
'	Else
'		Do While Not rst.Eof
'			sReturn= sReturn & "<option value="&rst("id")&">" & rst("classname") & "</option>" & VBCRLF
'			rst.Movenext
'		Loop
'		sReturn = "<option value=0>所有分类</option>" & VBCRLF & sReturn
		sReturn="<form name=selectsubmitform action='HomestayFamilyChina.asp' method=get><tr><td clospan=2>" & strType_city & str_usertype & "<Div id='line02' style='margin:0px;padding-top:6px;'>&nbsp;房屋结构&nbsp;<select name='houseframe' id='houseframe1' onchange='submit();' title='请选择您的房屋结构'>"&_
                            "<option value='' selected>室</option>"&_
							"<option value='1室'"
							If Instr(Request_houseframe,"1室")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">1室</option>"&_
                            "<option value='2室'"
							If Instr(Request_houseframe,"2室")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">2室</option>"&_
                            "<option value='3室'"
							If Instr(Request_houseframe,"3室")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">3室</option>"&_
                            "<option value='4室'"
							If Instr(Request_houseframe,"4室")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">4室</option>"&_
                            "<option value='5室'"
							If Instr(Request_houseframe,"5室")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">5室</option>"&_
                            "<option value='6室'"
							If Instr(Request_houseframe,"6室")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">6室</option>"&_
                          "</select>&nbsp;<select name='houseframe' id='houseframe2' onchange='submit();' title='请选择您的房屋结构'>"&_
						  "<option value='' selected>厅</option>"&_
                            "<option value='1厅'"
							If Instr(Request_houseframe,"1厅")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">1厅</option>"&_
                            "<option value='2厅'"
							If Instr(Request_houseframe,"2厅")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">2厅</option>"&_
                            "<option value='3厅'"
							If Instr(Request_houseframe,"3厅")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">3厅</option>"&_
                            "<option value='4厅'"
							If Instr(Request_houseframe,"4厅")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">4厅</option>"&_
                            "<option value='5厅'"
							If Instr(Request_houseframe,"5厅")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">5厅</option>"&_
                            "<option value='6厅'"
							If Instr(Request_houseframe,"6厅")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">6厅</option>"&_
                          "</select>&nbsp;<select name='houseframe' id='houseframe3' onchange='submit();' title='请选择您的房屋结构'>"&_
						  "<option value='' selected>卫</option>"&_
                            "<option value='1卫'"
							If Instr(Request_houseframe,"1卫")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">1卫</option>"&_
                            "<option value='2卫'"
							If Instr(Request_houseframe,"2卫")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">2卫</option>"&_
                            "<option value='3卫'"
							If Instr(Request_houseframe,"3卫")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">3卫</option>"&_
                            "<option value='4卫'"
							If Instr(Request_houseframe,"4卫")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">4卫</option>"&_
                            "<option value='5卫'"
							If Instr(Request_houseframe,"5卫")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">5卫</option>"&_
                            "<option value='6卫'"
							If Instr(Request_houseframe,"6卫")>0 Then sReturn=sReturn &" selected"
							sReturn=sReturn & ">6卫</option>"&_
                          "</select></Div> </tr></form>"
'	End If
'	rst.Close
'	Set rst=Nothing
	GetSysClassesSearchHome_start = sReturn
End Function

GetSysClassesSearchHome_start()
%>
123