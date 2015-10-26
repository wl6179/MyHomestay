<%
sub Inc_Admin_ShowClass_Option_clinic_self_Dept_Onlyclinic(ShowType,CurrentID)   '医院科室 ；给医生引用和其他引用的时候；
	if ShowType=0 then							'[2006-12-20   wl修改：]‘限制“仅门诊科室Only_clinic”的情况！’
	    response.write "<option value=0"
		if CurrentID=0 then response.write " selected"
		response.write ">无（作为一级栏目）</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	'sqlClass="Select * From oblog_logclass Where idType=" & t &" order by RootID,OrderID"
															't="0"   tName="文章
	sqlClass="Select * From clinic_self_Dept "&_
			"  Where  base_Hosp_ID="&base_Hosp_ID &_
			" And idType=0 "&_
			" And Only_clinic=0 "&_
			" order by RootID,OrderID"                '专门给其他引用！
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>请先添加栏目</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value=" & rsClass("classid") & ""    'rsClass("id")-------->rsClass("classid")wl...必须使用唯一值   '继续是用id而不改用classid了~wl
			if CurrentID>0 and rsClass("classid")=CurrentID then    'rsClass("id")-------->rsClass("classid")wl...必须使用唯一值   '继续是用id而不改用classid了~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "├&nbsp;"
						else
							strTemp=strTemp & "└&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "│"
						else
							strTemp=strTemp & "&nbsp;"
						end if
					end if
				next
			end if
			strTemp=strTemp & rsClass("classname")
			strTemp=strTemp & "</option>"
			response.write strTemp
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
end sub
%>
<%
sub Inc_Admin_ShowClass_Option_clinic_self_Dept(ShowType,CurrentID)   '医院科室 ；给医生引用和其他引用的时候；
	if ShowType=0 then							
	    response.write "<option value=0"
		if CurrentID=0 then response.write " selected"
		response.write ">无（作为一级栏目）</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	'sqlClass="Select * From oblog_logclass Where idType=" & t &" order by RootID,OrderID"
															't="0"   tName="文章
	sqlClass="Select * From clinic_self_Dept "&_
			"  Where  base_Hosp_ID="&base_Hosp_ID &_
			" And idType=0 "&_
			" order by RootID,OrderID"                '专门给其他引用！
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>请先添加栏目</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value=" & rsClass("classid") & ""    'rsClass("id")-------->rsClass("classid")wl...必须使用唯一值   '继续是用id而不改用classid了~wl
			if CurrentID>0 and rsClass("classid")=CurrentID then    'rsClass("id")-------->rsClass("classid")wl...必须使用唯一值   '继续是用id而不改用classid了~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "├&nbsp;"
						else
							strTemp=strTemp & "└&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "│"
						else
							strTemp=strTemp & "&nbsp;"
						end if
					end if
				next
			end if
			strTemp=strTemp & rsClass("classname")
			strTemp=strTemp & "</option>"
			response.write strTemp
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
end sub
sub Inc_Admin_ShowClass_Option_clinic_self_Dept_ForClinicSelfDept(ShowType,CurrentID)   '医院科室 ；专门给医院科室‘显示父级’用！
	if ShowType=0 then
	    response.write "<option value=0"
		if CurrentID=0 then response.write " selected"
		response.write ">无（作为一级栏目）</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	'sqlClass="Select * From oblog_logclass Where idType=" & t &" order by RootID,OrderID"
															't="0"   tName="文章
'	sqlClass="Select * From clinic_self_Dept "&_
'			"  Where  base_Hosp_ID="&base_Hosp_ID &_
'			" And idType=0 "&_
'			" order by RootID,OrderID"
	sqlClass="Select * From clinic_self_Dept "&_
			"  Where Only_clinic=0 And base_Hosp_ID="&base_Hosp_ID &_
			" And idType=0 "&_
			" order by RootID,OrderID"                  '专门给医院科室‘显示父级’用！
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>请先添加栏目</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value=" & rsClass("classid") & ""    'rsClass("id")-------->rsClass("classid")wl...必须使用唯一值   '继续是用id而不改用classid了~wl
			if CurrentID>0 and rsClass("classid")=CurrentID then    'rsClass("id")-------->rsClass("classid")wl...必须使用唯一值   '继续是用id而不改用classid了~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "├&nbsp;"
						else
							strTemp=strTemp & "└&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "│"
						else
							strTemp=strTemp & "&nbsp;"
						end if
					end if
				next
			end if
			strTemp=strTemp & rsClass("classname")
			strTemp=strTemp & "</option>"
			response.write strTemp
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
end sub


sub Inc_Admin_ShowClass_Option_clinic_HospDeployDept(ShowType,CurrentID)     '门诊部配置科室；
'Begin：wl 添加目的:非系统科室列表；列出的是医院自己设定的科室；
'End：  wl
	if ShowType=0 then
	    response.write "<option value='0'"
		if CurrentID=0 then response.write " selected"
		response.write ">无（作为一级栏目）</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	
	'Begin：wl 添加目的:非系统科室列表；列出的是医院自己设定的科室；[clinic_self_Dept]
	sqlClass="Select * From clinic_HospDeployDept where clinic_Hosp_ID=" & clinic_Hosp_ID & " order by RootID,OrderID"
	'End：  wl 
	
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>请先添加栏目</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value='" & rsClass("classid") & "'"   'rsClass("id")-------->rsClass("classid")wl...必须使用唯一值   '继续是用id而不改用classid了~wl
			if CurrentID>0 and rsClass("classid")=CurrentID then     'rsClass("id")-------->rsClass("classid")wl...必须使用唯一值   '继续是用id而不改用classid了~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "├&nbsp;"
						else
							strTemp=strTemp & "└&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "│"
						else
							strTemp=strTemp & "&nbsp;"
						end if
					end if
				next
			end if
			strTemp=strTemp & rsClass("classname")
			strTemp=strTemp & "</option>"
			response.write strTemp
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
end sub

%> 



<%
sub Inc_Admin_ShowClass_Option_clinic_self_Dept_02(ShowType,CurrentID)   '医院科ParentID室 ；针对ParentID的情况下，可能要用id的情况--------------------------------------------------------------------------------wl
	if ShowType=0 then
	    response.write "<option value=0"
		if CurrentID=0 then response.write " selected"
		response.write ">无（作为一级栏目）</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	'sqlClass="Select * From oblog_logclass Where idType=" & t &" order by RootID,OrderID"
															't="0"   tName="文章
	sqlClass="Select * From clinic_self_Dept "&_
			"  Where Only_clinic=0 And base_Hosp_ID="&base_Hosp_ID &_
			" And idType=0 "&_
			" order by RootID,OrderID"
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>请先添加栏目</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value=" & rsClass("classid") & ""    '此处仅仅针对ParentID来比较的情况下；   '继续是用id而不改用classid了~wl
			if CurrentID>0 and rsClass("id")=CurrentID then    'rsClass("id")-------->rsClass("classid")wl...必须使用唯一值   '继续是用id而不改用classid了~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "├&nbsp;"
						else
							strTemp=strTemp & "└&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "│"
						else
							strTemp=strTemp & "&nbsp;"
						end if
					end if
				next
			end if
			strTemp=strTemp & rsClass("classname")
			strTemp=strTemp & "</option>"
			response.write strTemp
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
end sub



sub Inc_Admin_ShowClass_Option_clinic_HospDeployDept_02(ShowType,CurrentID)     '门诊部配置科室；针对ParentID的情况下，可能要用id的情况--------------------------------------------------------------------------------wl
'Begin：wl 添加目的:非系统科室列表；列出的是医院自己设定的科室；
'End：  wl
	if ShowType=0 then
	    response.write "<option value='0'"
		if CurrentID=0 then response.write " selected"
		response.write ">无（作为一级栏目）</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	
	'Begin：wl 添加目的:非系统科室列表；列出的是医院自己设定的科室；[clinic_self_Dept]
	sqlClass="Select * From clinic_HospDeployDept where clinic_Hosp_ID=" & clinic_Hosp_ID & " order by RootID,OrderID"
	'End：  wl 
	
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>请先添加栏目</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value='" & rsClass("classid") & "'"    '此处仅仅针对ParentID来比较的情况下；   '继续是用id而不改用classid了~wl
			if CurrentID>0 and rsClass("id")=CurrentID then     'rsClass("id")-------->rsClass("classid")wl...必须使用唯一值   '继续是用id而不改用classid了~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "├&nbsp;"
						else
							strTemp=strTemp & "└&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "│"
						else
							strTemp=strTemp & "&nbsp;"
						end if
					end if
				next
			end if
			strTemp=strTemp & rsClass("classname")
			strTemp=strTemp & "</option>"
			response.write strTemp
			rsClass.movenext
		loop
	end if
	rsClass.close
	set rsClass=nothing
end sub

%> 


