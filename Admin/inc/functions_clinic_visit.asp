<%
'*********** 文件说明 *************
'
'     出诊、停诊信息的生成过程的定义文件
'
'*********** 文件说明 *************


'Begin: Thh: 获取医生在门诊科室的出诊信息
'一、在数据库表clinic_Visit中建唯一索引：同一科室、同一医生、同一星期、同一时间段的记录只有一条
'
'二、获取医生在某一个门诊科室的出诊信息:
'1. 先判断医生的出诊类别有那几种，按类型值的降序排列，即：4专科、3特需、2专家、1普通
'2. 对同一类出诊，都假定其出诊费用相同，获取出诊的费用
'3. 对同一类出诊，按时间段值的升序排列，即：1上午、2下午
'4. 分别循环每一类出诊的各个时间段，获取上下午时段出诊星期的字符串
'5. 比较上下午出诊星期的字符串是否相同，相同为全天，不同按上下午列出该出诊类别的出诊信息
'   注意排除空字符串和先过滤‘周’字
Sub UpdateClinicDeptDoctorVisit(lngDoctorID, lngClinicDeptID)
	
	Dim rsClinicType	'门诊类型的结果集
	Dim rsVisit			'门诊出诊记录 
	Dim sqlVisit
	
	'先判断医生的出诊类别有那几种
	Set rsClinicType = conn.Execute("Select  Distinct TypeNumber, TypeName " & _
										"From clinic_Visit " & _
										"Where Doctor_ID=" & lngDoctorID & " And clinicDept_ID=" & lngClinicDeptID & " " & _
										"Order By TypeNumber Desc")
	'获取医生的出诊记录
	Set rsVisit=Server.CreateObject("Adodb.Recordset")
	sqlVisit = "Select TypeName, TypeNumber, WeekName, WeekNumber, TimeName, TimeNumber, MoneyName, Doctor_ID, clinicDept_Name " & _
									"From clinic_Visit " & _
									"Where Doctor_ID=" & lngDoctorID & " And clinicDept_ID=" & lngClinicDeptID & " " & _
									"Order By TypeNumber Desc, TimeNumber, WeekNumber Asc"
	rsVisit.Open sqlVisit, conn,1,3	
	
	Dim strAllVisit	'出诊信息字符串
	
	'循环每一类出诊的各个时间段，获取上下午时段出诊星期的字符串		
	Do While Not rsClinicType.Eof
		
		Dim strWeekAm, strWeekPm, strTypeVisit, strMoney
		strWeekAm = ""
		strWeekPm = ""
		strTypeVisit = ""
		strMoney = ""
		
		'A. 获取当前出诊类别的上午的出诊记录集，以获取上午的出诊星期字符串
		rsVisit.Filter = "TimeNumber = 1 And TypeNumber = " & rsClinicType("TypeNumber")
		
		'A.1 获取金额
		If Not rsVisit.Eof Then
			strMoney = rsVisit("MoneyName")
		End If
		
		'A.2 获取周
		Do While Not rsVisit.Eof
			strWeekAm = strWeekAm & rsVisit("WeekName")

			rsVisit.MoveNext
		Loop
		
		'对出诊字符处理，加入逗号分隔符，从第二个字符开始替换，已经保证第一个字符为“周”
		If strWeekAm <> "" Then
			strWeekAm = "周" & Replace(strWeekAm, "周", "、", 2, -1, 1)
		End If
		
		'重新获取出诊信息的结果集，以获取下午的结果集
		rsVisit.Filter =""
		rsVisit.MoveFirst
		
		'B. 获取当前出诊类别的下午的出诊记录集，以获取下午的出诊星期字符串
		rsVisit.Filter = "TimeNumber = 2 And TypeNumber = " & rsClinicType("TypeNumber")
		
		'A.1 获取金额，因为上午可能没有出诊记录
		If Not rsVisit.Eof Then
			strMoney = rsVisit("MoneyName")
		End If
		
		'A.2 获取周
		Do While Not rsVisit.Eof
			strWeekPm = strWeekPm & rsVisit("WeekName")

			rsVisit.MoveNext
		Loop
		
		'对出诊字符处理，加入逗号分隔符，从第二个字符开始替换，已经保证第一个字符为“周”
		IF strWeekPm <> "" Then
			strWeekPm = "周" & Replace(strWeekPm, "周", "、", 2, -1, 1)
		End If
		
		'重新获取出诊信息的结果集，进入下一个出诊类别的循环
		rsVisit.Filter =""
		rsVisit.MoveFirst
		
		'C. 对上、下午的星期字符串进行处理
		'(1)上下午相同，且不为空字符串
		If strWeekAm = strWeekPm And strWeekAm <> "" Then
			strTypeVisit = strWeekAm & "全天"
		'(2)上下午不同，在判断是否为空字符串
		Else 
			'有上、下午的星期
			If strWeekAm <> "" And strWeekPm <> "" Then
				strTypeVisit = strWeekAm & "上午" & "；" & strWeekPm & "下午"
			Elseif strWeekAm <> "" Then
				strTypeVisit = strWeekAm & "上午"
			Elseif strWeekPm <> "" Then
				strTypeVisit = strWeekPm & "下午"
			End If
			
		End If
		
		'出诊信息不为空，在前加入出诊类别，在后加入该类别的费用，已经假定同一医生其在同一出诊类别上的费用相同
		If strTypeVisit <> "" Then
			strTypeVisit = rsClinicType("TypeName") & "：" & strTypeVisit 
			
			'费用不为空时加上费用，已经假定同一医生其在同一出诊类别上的费用相同，故可以直接使用出诊记录中的费用字段
			If strMoney <> "" Then
				strTypeVisit = strTypeVisit & "(" & strMoney  & ")"
			End If
		End If
		
		strAllVisit = strAllVisit & strTypeVisit & "　"
		
		rsClinicType.MoveNext
	Loop
		
	'更新医生在该门诊科室的出诊记录字段
	conn.Execute("Update clinic_ClinicDept_Doctor Set ClinicDept_Visit = '" & strAllVisit & "' " & _
						"Where Doctor_ID=" & lngDoctorID & " And clinicDept_ID=" & lngClinicDeptID)
	
	rsVisit.Close
	Set rsVisit = Nothing
	
	rsClinicType.Close
	Set rsClinicType = Nothing
	
End Sub


'Begin: Thh: 更新医生在门诊科室的停诊信息
'一、将停诊结束时间与当前时间比较：
'1. <0: 为正在出诊
'2. =0：为今日停诊
'3. =1：为明日停诊
'4. >1: 为停诊时间段：标明停诊的开始和结束时间，包含该时间点
'二、参数：
'1. lngDoctorID:     医生的主键ID
'2. lngClinicDeptID：门诊科室的主键ID
Sub UpdateClinicDeptDoctorStop(lngDoctorID, lngClinicDeptID)
	
	Dim rsStop			'门诊停诊记录 
	Dim sqlStop	
	Dim strStop			'停诊信息字符串
	strStop = ""
	
	'将停诊时间与当前时间比较，获取最晚的停诊时间，根据比较的天数来判断停诊情况
	sqlStop = "Select Top 1 onStop = DATEDIFF(day, getdate(), Stop_EndDate), " & _
							"stopDays = DATEDIFF(day, Stop_BeginDate, Stop_EndDate),  Stop_BeginDate, Stop_EndDate " &_
					"From clinic_Visit " &_
					"Where Doctor_ID=" & lngDoctorID & " And clinicDept_ID=" & lngClinicDeptID & " " & _
					"Order by onStop Desc" 
	Dim onStop		'停诊的状态，<0: 为正在出诊
	Dim stopDays	'停诊的时间
	
	Set rsStop = oblog.Execute(sqlStop)
	
	If Not (rsStop.Eof And rsStop.Bof) Then
		onStop   = rsStop("onStop")
		stopDays = rsStop("stopDays")
		
		'没有停诊
		If onStop < 0 Then
			strStop = ""
		'最新停诊时间为当天，并且停诊开始和结束时间相同
		Elseif onStop = 0 And stopDays = 0 Then
			strStop = Month(rsStop("Stop_EndDate")) & "." & Day(rsStop("Stop_EndDate")) & "日停诊"		'样式：12.30日停诊
		
		'最新停诊为当天或多天，且停诊据今天的时间在180天内
		Elseif onStop >= 0 And onStop < 180 Then
			If stopDays = 0 Then	'停诊一天
				strStop = Month(rsStop("Stop_EndDate")) & "." & Day(rsStop("Stop_EndDate")) & "日停诊"		'样式：12.30日停诊
			Else	
				strStop = Month(rsStop("Stop_BeginDate")) & "." & Day(rsStop("Stop_BeginDate")) & "-" & _
							Month(rsStop("Stop_EndDate")) & "." & Day(rsStop("Stop_EndDate")) & "日停诊"
			End If
			
		Elseif onStop > 0 And onStop > 180 Then
			strStop = "自" & FormatDateTime(rsStop("Stop_BeginDate"), 2) & "起长期停诊"
		
		Else 
			strStop = "停诊错误"
		End If
		
	End If
	
	'更新医生在该门诊科室的出诊记录字段
	conn.Execute("Update clinic_ClinicDept_Doctor Set ClinicDept_Stop = '" & strStop & "' " & _
						"Where Doctor_ID=" & lngDoctorID & " And clinicDept_ID=" & lngClinicDeptID)
	
	rsStop.Close
	Set rsStop = Nothing
	
End Sub
'End: Thh: 更新医生在门诊科室的停诊信息

'唐海华：添加：更新整个门诊部的医生停诊信息
'一、代码流程、步骤：
'1. 查找门诊部的所有门诊科室
'2. 对每个门诊科室，再查找该科室下的所有医生
'3. 对每个医生调用：UpdateClinicDeptDoctorStop(lngDoctorID, lngClinicDeptID) 更新该医生在该门诊科室的停诊记录
'二、参数：
'1. lngClinicHospID: 门诊部的主键ID
Sub UpdateClinicHospDoctorStop(lngClinicHospID)
	Dim rsDept, rsDoctor
	Dim sqlDept, sqlDoctor	
	
	'1. 查找门诊部下的所有门诊科室
	sqlDept = "Select classid, classname, depth, clinic_Hosp_ID " & _
					"From clinic_HospDeployDept " & _
					"Where clinic_Hosp_ID = " & lngClinicHospID
	Set rsDept = oblog.Execute(sqlDept)
	
	'2. 循环门诊部的所有门诊科室
	Do While Not rsDept.Eof
		
		'2.1 查找门诊科室下的所有医生
		sqlDoctor = "Select id, Doctor_ID, clinicDept_ID, clinicDept_Visit, clinicDept_Stop " & _
						"From clinic_ClinicDept_Doctor " & _
						"Where clinicDept_ID = " & rsDept("classid")
		Set rsDoctor = oblog.Execute(sqlDoctor)
			
		'2.2 循环门诊科室的所有医生，更新医生在该门诊科室的出诊信息
		Do While Not rsDoctor.Eof
			
			'3.1 调用更新医生出诊信息的过程：UpdateClinicDeptDoctorStop(lngDoctorID, lngClinicDeptID)
			Call UpdateClinicDeptDoctorStop(rsDoctor("Doctor_ID"), rsDept("classid"))
			
			'3.1 进入该科室的下一个医生
			rsDoctor.MoveNext
		Loop
		
		'2.3 进入该门诊部的下一个门诊科室
		rsDept.MoveNext
	Loop
	
	rsDoctor.Close
	Set rsDoctor = Nothing
	
	rsDept.Close
	Set rsDept = Nothing

End Sub


'唐海华：添加：更新整个门诊部的医生出诊信息
'一、代码流程、步骤：
'1. 查找门诊部的所有门诊科室
'2. 对每个门诊科室，再查找该科室下的所有医生
'3. 对每个医生调用：UpdateClinicDeptDoctorVisit(lngDoctorID, lngClinicDeptID) 更新该医生在该门诊科室的出诊记录
'二、参数：
'1. lngClinicHospID: 门诊部的主键ID
Sub UpdateClinicHospDoctorVisit(lngClinicHospID)
	Dim rsDept, rsDoctor
	Dim sqlDept, sqlDoctor	
	
	'1. 查找门诊部下的所有门诊科室
	sqlDept = "Select classid, classname, depth, clinic_Hosp_ID " & _
					"From clinic_HospDeployDept " & _
					"Where clinic_Hosp_ID = " & lngClinicHospID
	Set rsDept = oblog.Execute(sqlDept)
	
	'2. 循环门诊部的所有门诊科室
	Do While Not rsDept.Eof
		
		'2.1 查找门诊科室下的所有医生
		sqlDoctor = "Select id, Doctor_ID, clinicDept_ID, clinicDept_Visit, clinicDept_Stop " & _
						"From clinic_ClinicDept_Doctor " & _
						"Where clinicDept_ID = " & rsDept("classid")
		Set rsDoctor = oblog.Execute(sqlDoctor)
			
		'2.2 循环门诊科室的所有医生，更新医生在该门诊科室的出诊信息
		Do While Not rsDoctor.Eof
			
			'3.1 调用更新医生出诊信息的过程：UpdateClinicDeptDoctorStop(lngDoctorID, lngClinicDeptID)
			Call UpdateClinicDeptDoctorVisit(rsDoctor("Doctor_ID"), rsDept("classid"))
			
			'3.1 进入该科室的下一个医生
			rsDoctor.MoveNext
		Loop
		
		'2.3 进入该门诊部的下一个门诊科室
		rsDept.MoveNext
	Loop
	
	rsDoctor.Close
	Set rsDoctor = Nothing
	
	rsDept.Close
	Set rsDept = Nothing

End Sub


'唐海华添加：设置医生停诊的某天的日期
'一、代码流程：
'1. 将停诊的起始日期设置为哪天
'2. 调用：UpdateClinicDeptDoctorStop() 来更新医生在门诊科室的停诊信息
'二、参数：
'1. id:              出诊记录的主键
'1. lngDoctorID:     医生的主键ID
'2. lngClinicDeptID：门诊科室的主键ID
'3. dateStop：       停诊日期 为日期类型
Sub StopClinicDeptDoctorOneDay(id, lngDoctorID, lngClinicDeptID, dateStop)
	'A 检查停诊的日期是否合法
	If Not IsDate(dateStop) Then
		'不为日期类型则退出
		Response.Write("参数部位日期类型")
		Exit Sub
	End If
	
	'B 更新停诊的起始日期
	conn.Execute("Update clinic_Visit Set Stop_BeginDate = '" & dateStop & "', Stop_EndDate = '" & dateStop & "' " & _
						"Where id=" & id)
	
	'C 更新医生的停诊信息
	Call UpdateClinicDeptDoctorStop(lngDoctorID, lngClinicDeptID)

End Sub

%>



<%Sub show_selectdate(addtime)              
 dim y,m,d,h,mi,s,ttime
 if addtime="" then ttime=ServerDate(now()) else ttime=addtime
 response.Write("<select name=selecty id=selecty>")
 for  y=2000 to 2010
  if year(ttime)=y then
   response.Write "<option value="&y&" selected>"&y&"年</option>"
  else
   response.Write "<option value="&y&">"&y&"年</option>"
  end if
 next
 response.Write "</select><select name=selectm id=selectm >"
 for m=1 to 12
    if month(ttime)=m then
     response.Write "<option value="&m&" selected>"&m&"月</option>"
  else
     response.Write "<option value="&m&">"&m&"月</option>"
  end if
 next
 response.Write("</select><select name=selectd id=selectd >")
 for d=1 to 31
  if day(ttime)=d then
     response.Write "<option value="&d&" selected>"&d&"日</option>"
  else
     response.Write "<option value="&d&">"&d&"日</option>"
  end if
 next
 response.Write("</select><select name=selecth id=selecth>")            
 for h=0 to 23
    if hour(ttime)=h then
     response.Write "<option value="&h&" selected>"&h&"时</option>"
  else
     response.Write "<option value="&h&">"&h&"时</option>"
  end if
 next
 response.Write("</select><select name=selectmi id=selectmi>")      
 for mi=0 to 59
    if minute(ttime)=mi then
     response.Write "<option value="&mi&" selected>"&mi&"分</option>"
  else
     response.Write "<option value="&mi&">"&mi&"分</option>"
  end if
 next
 response.Write("</select><select name=selects id=selects>")
 for s=0 to 59
    if second(ttime)=s then
     response.Write "<option value="&s&" selected>"&s&"秒</option>"
  else
     response.Write "<option value="&s&">"&s&"秒</option>"
  end if
 next
 response.Write("</select>")  
End Sub
%>