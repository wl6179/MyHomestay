<%
sub Inc_Admin_ShowClass_Option_clinic_self_Dept_Onlyclinic(ShowType,CurrentID)   'ҽԺ���� ����ҽ�����ú��������õ�ʱ��
	if ShowType=0 then							'[2006-12-20   wl�޸ģ�]�����ơ����������Only_clinic�����������
	    response.write "<option value=0"
		if CurrentID=0 then response.write " selected"
		response.write ">�ޣ���Ϊһ����Ŀ��</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	'sqlClass="Select * From oblog_logclass Where idType=" & t &" order by RootID,OrderID"
															't="0"   tName="����
	sqlClass="Select * From clinic_self_Dept "&_
			"  Where  base_Hosp_ID="&base_Hosp_ID &_
			" And idType=0 "&_
			" And Only_clinic=0 "&_
			" order by RootID,OrderID"                'ר�Ÿ��������ã�
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>���������Ŀ</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value=" & rsClass("classid") & ""    'rsClass("id")-------->rsClass("classid")wl...����ʹ��Ψһֵ   '��������id��������classid��~wl
			if CurrentID>0 and rsClass("classid")=CurrentID then    'rsClass("id")-------->rsClass("classid")wl...����ʹ��Ψһֵ   '��������id��������classid��~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "��&nbsp;"
						else
							strTemp=strTemp & "��&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "��"
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
sub Inc_Admin_ShowClass_Option_clinic_self_Dept(ShowType,CurrentID)   'ҽԺ���� ����ҽ�����ú��������õ�ʱ��
	if ShowType=0 then							
	    response.write "<option value=0"
		if CurrentID=0 then response.write " selected"
		response.write ">�ޣ���Ϊһ����Ŀ��</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	'sqlClass="Select * From oblog_logclass Where idType=" & t &" order by RootID,OrderID"
															't="0"   tName="����
	sqlClass="Select * From clinic_self_Dept "&_
			"  Where  base_Hosp_ID="&base_Hosp_ID &_
			" And idType=0 "&_
			" order by RootID,OrderID"                'ר�Ÿ��������ã�
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>���������Ŀ</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value=" & rsClass("classid") & ""    'rsClass("id")-------->rsClass("classid")wl...����ʹ��Ψһֵ   '��������id��������classid��~wl
			if CurrentID>0 and rsClass("classid")=CurrentID then    'rsClass("id")-------->rsClass("classid")wl...����ʹ��Ψһֵ   '��������id��������classid��~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "��&nbsp;"
						else
							strTemp=strTemp & "��&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "��"
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
sub Inc_Admin_ShowClass_Option_clinic_self_Dept_ForClinicSelfDept(ShowType,CurrentID)   'ҽԺ���� ��ר�Ÿ�ҽԺ���ҡ���ʾ�������ã�
	if ShowType=0 then
	    response.write "<option value=0"
		if CurrentID=0 then response.write " selected"
		response.write ">�ޣ���Ϊһ����Ŀ��</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	'sqlClass="Select * From oblog_logclass Where idType=" & t &" order by RootID,OrderID"
															't="0"   tName="����
'	sqlClass="Select * From clinic_self_Dept "&_
'			"  Where  base_Hosp_ID="&base_Hosp_ID &_
'			" And idType=0 "&_
'			" order by RootID,OrderID"
	sqlClass="Select * From clinic_self_Dept "&_
			"  Where Only_clinic=0 And base_Hosp_ID="&base_Hosp_ID &_
			" And idType=0 "&_
			" order by RootID,OrderID"                  'ר�Ÿ�ҽԺ���ҡ���ʾ�������ã�
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>���������Ŀ</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value=" & rsClass("classid") & ""    'rsClass("id")-------->rsClass("classid")wl...����ʹ��Ψһֵ   '��������id��������classid��~wl
			if CurrentID>0 and rsClass("classid")=CurrentID then    'rsClass("id")-------->rsClass("classid")wl...����ʹ��Ψһֵ   '��������id��������classid��~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "��&nbsp;"
						else
							strTemp=strTemp & "��&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "��"
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


sub Inc_Admin_ShowClass_Option_clinic_HospDeployDept(ShowType,CurrentID)     '���ﲿ���ÿ��ң�
'Begin��wl ���Ŀ��:��ϵͳ�����б��г�����ҽԺ�Լ��趨�Ŀ��ң�
'End��  wl
	if ShowType=0 then
	    response.write "<option value='0'"
		if CurrentID=0 then response.write " selected"
		response.write ">�ޣ���Ϊһ����Ŀ��</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	
	'Begin��wl ���Ŀ��:��ϵͳ�����б��г�����ҽԺ�Լ��趨�Ŀ��ң�[clinic_self_Dept]
	sqlClass="Select * From clinic_HospDeployDept where clinic_Hosp_ID=" & clinic_Hosp_ID & " order by RootID,OrderID"
	'End��  wl 
	
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>���������Ŀ</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value='" & rsClass("classid") & "'"   'rsClass("id")-------->rsClass("classid")wl...����ʹ��Ψһֵ   '��������id��������classid��~wl
			if CurrentID>0 and rsClass("classid")=CurrentID then     'rsClass("id")-------->rsClass("classid")wl...����ʹ��Ψһֵ   '��������id��������classid��~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "��&nbsp;"
						else
							strTemp=strTemp & "��&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "��"
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
sub Inc_Admin_ShowClass_Option_clinic_self_Dept_02(ShowType,CurrentID)   'ҽԺ��ParentID�� �����ParentID������£�����Ҫ��id�����--------------------------------------------------------------------------------wl
	if ShowType=0 then
	    response.write "<option value=0"
		if CurrentID=0 then response.write " selected"
		response.write ">�ޣ���Ϊһ����Ŀ��</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	'sqlClass="Select * From oblog_logclass Where idType=" & t &" order by RootID,OrderID"
															't="0"   tName="����
	sqlClass="Select * From clinic_self_Dept "&_
			"  Where Only_clinic=0 And base_Hosp_ID="&base_Hosp_ID &_
			" And idType=0 "&_
			" order by RootID,OrderID"
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>���������Ŀ</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value=" & rsClass("classid") & ""    '�˴��������ParentID���Ƚϵ�����£�   '��������id��������classid��~wl
			if CurrentID>0 and rsClass("id")=CurrentID then    'rsClass("id")-------->rsClass("classid")wl...����ʹ��Ψһֵ   '��������id��������classid��~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "��&nbsp;"
						else
							strTemp=strTemp & "��&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "��"
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



sub Inc_Admin_ShowClass_Option_clinic_HospDeployDept_02(ShowType,CurrentID)     '���ﲿ���ÿ��ң����ParentID������£�����Ҫ��id�����--------------------------------------------------------------------------------wl
'Begin��wl ���Ŀ��:��ϵͳ�����б��г�����ҽԺ�Լ��趨�Ŀ��ң�
'End��  wl
	if ShowType=0 then
	    response.write "<option value='0'"
		if CurrentID=0 then response.write " selected"
		response.write ">�ޣ���Ϊһ����Ŀ��</option>"
	end if
	dim rsClass,sqlClass,strTemp,tmpDepth,i
	dim arrShowLine(20)
	for i=0 to ubound(arrShowLine)
		arrShowLine(i)=False
	next
	
	'Begin��wl ���Ŀ��:��ϵͳ�����б��г�����ҽԺ�Լ��趨�Ŀ��ң�[clinic_self_Dept]
	sqlClass="Select * From clinic_HospDeployDept where clinic_Hosp_ID=" & clinic_Hosp_ID & " order by RootID,OrderID"
	'End��  wl 
	
	set rsClass=conn.execute(sqlClass)
	if rsClass.bof and rsClass.eof then 
		response.write "<option value=''>���������Ŀ</option>"
	else
		do while not rsClass.eof
			tmpDepth=rsClass("Depth")
			if rsClass("NextID")>0 then
				arrShowLine(tmpDepth)=True
			else
				arrShowLine(tmpDepth)=False
			end if
				strTemp="<option value='" & rsClass("classid") & "'"    '�˴��������ParentID���Ƚϵ�����£�   '��������id��������classid��~wl
			if CurrentID>0 and rsClass("id")=CurrentID then     'rsClass("id")-------->rsClass("classid")wl...����ʹ��Ψһֵ   '��������id��������classid��~wl
				 strTemp=strTemp & " selected"
			end if
			strTemp=strTemp & ">"
			
			if tmpDepth>0 then
				for i=1 to tmpDepth
					strTemp=strTemp & "&nbsp;&nbsp;"
					if i=tmpDepth then
						if rsClass("NextID")>0 then
							strTemp=strTemp & "��&nbsp;"
						else
							strTemp=strTemp & "��&nbsp;"
						end if
					else
						if arrShowLine(i)=True then
							strTemp=strTemp & "��"
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


