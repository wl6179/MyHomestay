
<%
'���ļ�Ϊ�Ƹ�֧ͨ�������ļ�

'�������������Ƹ�ͨ���Կ���   0 �رղ���    1 �������ԡ������������

spid		= "12***701"							' �����滻Ϊ����ʵ���̻���
sp_key		= "***"	' sp_key��32λ�̻���Կ, ���滻Ϊ����ʵ����Կ
tenpay_dir	= "/tenpay_bargainor"					' WL ����Ŀ¼��

'�������������Ƹ�֧ͨ��������������������
domain				="http://www.myhomestay.com.cn"									'�̻���վ����


site_name			="�ҵ�ס������ȫ֧��"											'�̻���վ����
imgtitle			="����ѡ��ȫ�ĲƸ�ͨ����֧����ʽ" 							'ͼƬ˵��
imgsrc				=tenpay_dir&"/image/tenpay_buy.gif"								'ͼƬ��ַ

'mch_returl			= "http://www.myhomestay.com.cn/tenpay_myhomestay.com.cn/tenpay_notify.asp"			'֪ͨURL
'show_url			= "http://www.myhomestay.com.cn/tenpay_myhomestay.com.cn/tenpay_show.asp"			'����URL





' WL�Լ���˾�����̻������ź��� �� ԭ����ȡ������ ʱ�䣬��ʽMMDDHHMMSS ��10λ���ȣ�
Function CFTGetServerDate_BillNo 
	Dim strTmp, iYear,iMonth,iDate 
	'//iYear = Year(Date) 
	'//iMonth = Month(Date) 
	'//iDate = Day(Date) 
	
	iHour = Hour(Now)
	iMinute = Minute(Now)
	iSecond = Second(Now)
	
	'//strTmp = CStr(iYear)
	strTmp = Right(SESSION.SessionID, 4)		'WL
	
'//	If iMonth < 10 Then 
'//		strTmp = strTmp & "0" & Cstr(iMonth)
'//	Else 
'//		strTmp = strTmp & Cstr(iMonth)
'//	End If
'//	
'//	If iDate < 10 Then 
'//		strTmp = strTmp & "0" & Cstr(iDate) 
'//	Else 
'//		strTmp = strTmp & Cstr(iDate) 
'//	End If
'//	
	
	
	If iHour < 10 Then 
		strTmp = strTmp & "0" & Cstr(iHour) 
	Else 
		strTmp = strTmp & Cstr(iHour) 
	End If
	
	If iMinute < 10 Then 
		strTmp = strTmp & "0" & Cstr(iMinute) 
	Else 
		strTmp = strTmp & Cstr(iMinute) 
	End If
	
	If iSecond < 10 Then 
		strTmp = strTmp & "0" & Cstr(iSecond) 
	Else 
		strTmp = strTmp & Cstr(iSecond) 
	End If
	
	CFTGetServerDate_BillNo = strTmp 
End Function

%>