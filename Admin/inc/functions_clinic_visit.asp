<%
'*********** �ļ�˵�� *************
'
'     ���ͣ����Ϣ�����ɹ��̵Ķ����ļ�
'
'*********** �ļ�˵�� *************


'Begin: Thh: ��ȡҽ����������ҵĳ�����Ϣ
'һ�������ݿ��clinic_Visit�н�Ψһ������ͬһ���ҡ�ͬһҽ����ͬһ���ڡ�ͬһʱ��εļ�¼ֻ��һ��
'
'������ȡҽ����ĳһ��������ҵĳ�����Ϣ:
'1. ���ж�ҽ���ĳ���������Ǽ��֣�������ֵ�Ľ������У�����4ר�ơ�3���衢2ר�ҡ�1��ͨ
'2. ��ͬһ�������ٶ�����������ͬ����ȡ����ķ���
'3. ��ͬһ������ʱ���ֵ���������У�����1���硢2����
'4. �ֱ�ѭ��ÿһ�����ĸ���ʱ��Σ���ȡ������ʱ�γ������ڵ��ַ���
'5. �Ƚ�������������ڵ��ַ����Ƿ���ͬ����ͬΪȫ�죬��ͬ���������г��ó������ĳ�����Ϣ
'   ע���ų����ַ������ȹ��ˡ��ܡ���
Sub UpdateClinicDeptDoctorVisit(lngDoctorID, lngClinicDeptID)
	
	Dim rsClinicType	'�������͵Ľ����
	Dim rsVisit			'��������¼ 
	Dim sqlVisit
	
	'���ж�ҽ���ĳ���������Ǽ���
	Set rsClinicType = conn.Execute("Select  Distinct TypeNumber, TypeName " & _
										"From clinic_Visit " & _
										"Where Doctor_ID=" & lngDoctorID & " And clinicDept_ID=" & lngClinicDeptID & " " & _
										"Order By TypeNumber Desc")
	'��ȡҽ���ĳ����¼
	Set rsVisit=Server.CreateObject("Adodb.Recordset")
	sqlVisit = "Select TypeName, TypeNumber, WeekName, WeekNumber, TimeName, TimeNumber, MoneyName, Doctor_ID, clinicDept_Name " & _
									"From clinic_Visit " & _
									"Where Doctor_ID=" & lngDoctorID & " And clinicDept_ID=" & lngClinicDeptID & " " & _
									"Order By TypeNumber Desc, TimeNumber, WeekNumber Asc"
	rsVisit.Open sqlVisit, conn,1,3	
	
	Dim strAllVisit	'������Ϣ�ַ���
	
	'ѭ��ÿһ�����ĸ���ʱ��Σ���ȡ������ʱ�γ������ڵ��ַ���		
	Do While Not rsClinicType.Eof
		
		Dim strWeekAm, strWeekPm, strTypeVisit, strMoney
		strWeekAm = ""
		strWeekPm = ""
		strTypeVisit = ""
		strMoney = ""
		
		'A. ��ȡ��ǰ������������ĳ����¼�����Ի�ȡ����ĳ��������ַ���
		rsVisit.Filter = "TimeNumber = 1 And TypeNumber = " & rsClinicType("TypeNumber")
		
		'A.1 ��ȡ���
		If Not rsVisit.Eof Then
			strMoney = rsVisit("MoneyName")
		End If
		
		'A.2 ��ȡ��
		Do While Not rsVisit.Eof
			strWeekAm = strWeekAm & rsVisit("WeekName")

			rsVisit.MoveNext
		Loop
		
		'�Գ����ַ��������붺�ŷָ������ӵڶ����ַ���ʼ�滻���Ѿ���֤��һ���ַ�Ϊ���ܡ�
		If strWeekAm <> "" Then
			strWeekAm = "��" & Replace(strWeekAm, "��", "��", 2, -1, 1)
		End If
		
		'���»�ȡ������Ϣ�Ľ�������Ի�ȡ����Ľ����
		rsVisit.Filter =""
		rsVisit.MoveFirst
		
		'B. ��ȡ��ǰ������������ĳ����¼�����Ի�ȡ����ĳ��������ַ���
		rsVisit.Filter = "TimeNumber = 2 And TypeNumber = " & rsClinicType("TypeNumber")
		
		'A.1 ��ȡ����Ϊ�������û�г����¼
		If Not rsVisit.Eof Then
			strMoney = rsVisit("MoneyName")
		End If
		
		'A.2 ��ȡ��
		Do While Not rsVisit.Eof
			strWeekPm = strWeekPm & rsVisit("WeekName")

			rsVisit.MoveNext
		Loop
		
		'�Գ����ַ��������붺�ŷָ������ӵڶ����ַ���ʼ�滻���Ѿ���֤��һ���ַ�Ϊ���ܡ�
		IF strWeekPm <> "" Then
			strWeekPm = "��" & Replace(strWeekPm, "��", "��", 2, -1, 1)
		End If
		
		'���»�ȡ������Ϣ�Ľ������������һ����������ѭ��
		rsVisit.Filter =""
		rsVisit.MoveFirst
		
		'C. ���ϡ�����������ַ������д���
		'(1)��������ͬ���Ҳ�Ϊ���ַ���
		If strWeekAm = strWeekPm And strWeekAm <> "" Then
			strTypeVisit = strWeekAm & "ȫ��"
		'(2)�����粻ͬ�����ж��Ƿ�Ϊ���ַ���
		Else 
			'���ϡ����������
			If strWeekAm <> "" And strWeekPm <> "" Then
				strTypeVisit = strWeekAm & "����" & "��" & strWeekPm & "����"
			Elseif strWeekAm <> "" Then
				strTypeVisit = strWeekAm & "����"
			Elseif strWeekPm <> "" Then
				strTypeVisit = strWeekPm & "����"
			End If
			
		End If
		
		'������Ϣ��Ϊ�գ���ǰ�����������ں��������ķ��ã��Ѿ��ٶ�ͬһҽ������ͬһ��������ϵķ�����ͬ
		If strTypeVisit <> "" Then
			strTypeVisit = rsClinicType("TypeName") & "��" & strTypeVisit 
			
			'���ò�Ϊ��ʱ���Ϸ��ã��Ѿ��ٶ�ͬһҽ������ͬһ��������ϵķ�����ͬ���ʿ���ֱ��ʹ�ó����¼�еķ����ֶ�
			If strMoney <> "" Then
				strTypeVisit = strTypeVisit & "(" & strMoney  & ")"
			End If
		End If
		
		strAllVisit = strAllVisit & strTypeVisit & "��"
		
		rsClinicType.MoveNext
	Loop
		
	'����ҽ���ڸ�������ҵĳ����¼�ֶ�
	conn.Execute("Update clinic_ClinicDept_Doctor Set ClinicDept_Visit = '" & strAllVisit & "' " & _
						"Where Doctor_ID=" & lngDoctorID & " And clinicDept_ID=" & lngClinicDeptID)
	
	rsVisit.Close
	Set rsVisit = Nothing
	
	rsClinicType.Close
	Set rsClinicType = Nothing
	
End Sub


'Begin: Thh: ����ҽ����������ҵ�ͣ����Ϣ
'һ����ͣ�����ʱ���뵱ǰʱ��Ƚϣ�
'1. <0: Ϊ���ڳ���
'2. =0��Ϊ����ͣ��
'3. =1��Ϊ����ͣ��
'4. >1: Ϊͣ��ʱ��Σ�����ͣ��Ŀ�ʼ�ͽ���ʱ�䣬������ʱ���
'����������
'1. lngDoctorID:     ҽ��������ID
'2. lngClinicDeptID��������ҵ�����ID
Sub UpdateClinicDeptDoctorStop(lngDoctorID, lngClinicDeptID)
	
	Dim rsStop			'����ͣ���¼ 
	Dim sqlStop	
	Dim strStop			'ͣ����Ϣ�ַ���
	strStop = ""
	
	'��ͣ��ʱ���뵱ǰʱ��Ƚϣ���ȡ�����ͣ��ʱ�䣬���ݱȽϵ��������ж�ͣ�����
	sqlStop = "Select Top 1 onStop = DATEDIFF(day, getdate(), Stop_EndDate), " & _
							"stopDays = DATEDIFF(day, Stop_BeginDate, Stop_EndDate),  Stop_BeginDate, Stop_EndDate " &_
					"From clinic_Visit " &_
					"Where Doctor_ID=" & lngDoctorID & " And clinicDept_ID=" & lngClinicDeptID & " " & _
					"Order by onStop Desc" 
	Dim onStop		'ͣ���״̬��<0: Ϊ���ڳ���
	Dim stopDays	'ͣ���ʱ��
	
	Set rsStop = oblog.Execute(sqlStop)
	
	If Not (rsStop.Eof And rsStop.Bof) Then
		onStop   = rsStop("onStop")
		stopDays = rsStop("stopDays")
		
		'û��ͣ��
		If onStop < 0 Then
			strStop = ""
		'����ͣ��ʱ��Ϊ���죬����ͣ�￪ʼ�ͽ���ʱ����ͬ
		Elseif onStop = 0 And stopDays = 0 Then
			strStop = Month(rsStop("Stop_EndDate")) & "." & Day(rsStop("Stop_EndDate")) & "��ͣ��"		'��ʽ��12.30��ͣ��
		
		'����ͣ��Ϊ�������죬��ͣ��ݽ����ʱ����180����
		Elseif onStop >= 0 And onStop < 180 Then
			If stopDays = 0 Then	'ͣ��һ��
				strStop = Month(rsStop("Stop_EndDate")) & "." & Day(rsStop("Stop_EndDate")) & "��ͣ��"		'��ʽ��12.30��ͣ��
			Else	
				strStop = Month(rsStop("Stop_BeginDate")) & "." & Day(rsStop("Stop_BeginDate")) & "-" & _
							Month(rsStop("Stop_EndDate")) & "." & Day(rsStop("Stop_EndDate")) & "��ͣ��"
			End If
			
		Elseif onStop > 0 And onStop > 180 Then
			strStop = "��" & FormatDateTime(rsStop("Stop_BeginDate"), 2) & "����ͣ��"
		
		Else 
			strStop = "ͣ�����"
		End If
		
	End If
	
	'����ҽ���ڸ�������ҵĳ����¼�ֶ�
	conn.Execute("Update clinic_ClinicDept_Doctor Set ClinicDept_Stop = '" & strStop & "' " & _
						"Where Doctor_ID=" & lngDoctorID & " And clinicDept_ID=" & lngClinicDeptID)
	
	rsStop.Close
	Set rsStop = Nothing
	
End Sub
'End: Thh: ����ҽ����������ҵ�ͣ����Ϣ

'�ƺ�������ӣ������������ﲿ��ҽ��ͣ����Ϣ
'һ���������̡����裺
'1. �������ﲿ�������������
'2. ��ÿ��������ң��ٲ��Ҹÿ����µ�����ҽ��
'3. ��ÿ��ҽ�����ã�UpdateClinicDeptDoctorStop(lngDoctorID, lngClinicDeptID) ���¸�ҽ���ڸ�������ҵ�ͣ���¼
'����������
'1. lngClinicHospID: ���ﲿ������ID
Sub UpdateClinicHospDoctorStop(lngClinicHospID)
	Dim rsDept, rsDoctor
	Dim sqlDept, sqlDoctor	
	
	'1. �������ﲿ�µ������������
	sqlDept = "Select classid, classname, depth, clinic_Hosp_ID " & _
					"From clinic_HospDeployDept " & _
					"Where clinic_Hosp_ID = " & lngClinicHospID
	Set rsDept = oblog.Execute(sqlDept)
	
	'2. ѭ�����ﲿ�������������
	Do While Not rsDept.Eof
		
		'2.1 ������������µ�����ҽ��
		sqlDoctor = "Select id, Doctor_ID, clinicDept_ID, clinicDept_Visit, clinicDept_Stop " & _
						"From clinic_ClinicDept_Doctor " & _
						"Where clinicDept_ID = " & rsDept("classid")
		Set rsDoctor = oblog.Execute(sqlDoctor)
			
		'2.2 ѭ��������ҵ�����ҽ��������ҽ���ڸ�������ҵĳ�����Ϣ
		Do While Not rsDoctor.Eof
			
			'3.1 ���ø���ҽ��������Ϣ�Ĺ��̣�UpdateClinicDeptDoctorStop(lngDoctorID, lngClinicDeptID)
			Call UpdateClinicDeptDoctorStop(rsDoctor("Doctor_ID"), rsDept("classid"))
			
			'3.1 ����ÿ��ҵ���һ��ҽ��
			rsDoctor.MoveNext
		Loop
		
		'2.3 ��������ﲿ����һ���������
		rsDept.MoveNext
	Loop
	
	rsDoctor.Close
	Set rsDoctor = Nothing
	
	rsDept.Close
	Set rsDept = Nothing

End Sub


'�ƺ�������ӣ������������ﲿ��ҽ��������Ϣ
'һ���������̡����裺
'1. �������ﲿ�������������
'2. ��ÿ��������ң��ٲ��Ҹÿ����µ�����ҽ��
'3. ��ÿ��ҽ�����ã�UpdateClinicDeptDoctorVisit(lngDoctorID, lngClinicDeptID) ���¸�ҽ���ڸ�������ҵĳ����¼
'����������
'1. lngClinicHospID: ���ﲿ������ID
Sub UpdateClinicHospDoctorVisit(lngClinicHospID)
	Dim rsDept, rsDoctor
	Dim sqlDept, sqlDoctor	
	
	'1. �������ﲿ�µ������������
	sqlDept = "Select classid, classname, depth, clinic_Hosp_ID " & _
					"From clinic_HospDeployDept " & _
					"Where clinic_Hosp_ID = " & lngClinicHospID
	Set rsDept = oblog.Execute(sqlDept)
	
	'2. ѭ�����ﲿ�������������
	Do While Not rsDept.Eof
		
		'2.1 ������������µ�����ҽ��
		sqlDoctor = "Select id, Doctor_ID, clinicDept_ID, clinicDept_Visit, clinicDept_Stop " & _
						"From clinic_ClinicDept_Doctor " & _
						"Where clinicDept_ID = " & rsDept("classid")
		Set rsDoctor = oblog.Execute(sqlDoctor)
			
		'2.2 ѭ��������ҵ�����ҽ��������ҽ���ڸ�������ҵĳ�����Ϣ
		Do While Not rsDoctor.Eof
			
			'3.1 ���ø���ҽ��������Ϣ�Ĺ��̣�UpdateClinicDeptDoctorStop(lngDoctorID, lngClinicDeptID)
			Call UpdateClinicDeptDoctorVisit(rsDoctor("Doctor_ID"), rsDept("classid"))
			
			'3.1 ����ÿ��ҵ���һ��ҽ��
			rsDoctor.MoveNext
		Loop
		
		'2.3 ��������ﲿ����һ���������
		rsDept.MoveNext
	Loop
	
	rsDoctor.Close
	Set rsDoctor = Nothing
	
	rsDept.Close
	Set rsDept = Nothing

End Sub


'�ƺ�����ӣ�����ҽ��ͣ���ĳ�������
'һ���������̣�
'1. ��ͣ�����ʼ��������Ϊ����
'2. ���ã�UpdateClinicDeptDoctorStop() ������ҽ����������ҵ�ͣ����Ϣ
'����������
'1. id:              �����¼������
'1. lngDoctorID:     ҽ��������ID
'2. lngClinicDeptID��������ҵ�����ID
'3. dateStop��       ͣ������ Ϊ��������
Sub StopClinicDeptDoctorOneDay(id, lngDoctorID, lngClinicDeptID, dateStop)
	'A ���ͣ��������Ƿ�Ϸ�
	If Not IsDate(dateStop) Then
		'��Ϊ�����������˳�
		Response.Write("������λ��������")
		Exit Sub
	End If
	
	'B ����ͣ�����ʼ����
	conn.Execute("Update clinic_Visit Set Stop_BeginDate = '" & dateStop & "', Stop_EndDate = '" & dateStop & "' " & _
						"Where id=" & id)
	
	'C ����ҽ����ͣ����Ϣ
	Call UpdateClinicDeptDoctorStop(lngDoctorID, lngClinicDeptID)

End Sub

%>



<%Sub show_selectdate(addtime)              
 dim y,m,d,h,mi,s,ttime
 if addtime="" then ttime=ServerDate(now()) else ttime=addtime
 response.Write("<select name=selecty id=selecty>")
 for  y=2000 to 2010
  if year(ttime)=y then
   response.Write "<option value="&y&" selected>"&y&"��</option>"
  else
   response.Write "<option value="&y&">"&y&"��</option>"
  end if
 next
 response.Write "</select><select name=selectm id=selectm >"
 for m=1 to 12
    if month(ttime)=m then
     response.Write "<option value="&m&" selected>"&m&"��</option>"
  else
     response.Write "<option value="&m&">"&m&"��</option>"
  end if
 next
 response.Write("</select><select name=selectd id=selectd >")
 for d=1 to 31
  if day(ttime)=d then
     response.Write "<option value="&d&" selected>"&d&"��</option>"
  else
     response.Write "<option value="&d&">"&d&"��</option>"
  end if
 next
 response.Write("</select><select name=selecth id=selecth>")            
 for h=0 to 23
    if hour(ttime)=h then
     response.Write "<option value="&h&" selected>"&h&"ʱ</option>"
  else
     response.Write "<option value="&h&">"&h&"ʱ</option>"
  end if
 next
 response.Write("</select><select name=selectmi id=selectmi>")      
 for mi=0 to 59
    if minute(ttime)=mi then
     response.Write "<option value="&mi&" selected>"&mi&"��</option>"
  else
     response.Write "<option value="&mi&">"&mi&"��</option>"
  end if
 next
 response.Write("</select><select name=selects id=selects>")
 for s=0 to 59
    if second(ttime)=s then
     response.Write "<option value="&s&" selected>"&s&"��</option>"
  else
     response.Write "<option value="&s&">"&s&"��</option>"
  end if
 next
 response.Write("</select>")  
End Sub
%>