<!-- #include file="inc/inc_syssite.asp" -->
<!-- #include File="inc/class_upfile.asp" -->
<!-- #include File="inc/class_blog.asp" -->
	<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<title>MyHomestay �ļ��ϴ�</title>
	<link href="css/back-ctrl.css" rel="stylesheet" type="text/css" />
	</head>
	<body>

<%
if not oblog.checkuserlogined() then
	response.Write("��½������ϴ��ļ�")
	response.End()
end if
Dim tMode
tMode=Request("tMode")'0log 1pho;
dim freesize,onesize,maxsize,enupload,upfiletype,re,ubb,isphoto
re=request.QueryString("re")
ubb=request.QueryString("ubb")
isphoto=cint(request.QueryString("isphoto"))
select case oblog.logined_ulevel
	case 6'WL�й���ͥ�ϴ�������
	enupload=oblog.setup(33,0)
	upfiletype=oblog.setup(34,0)
	onesize=oblog.setup(35,0)
	maxsize=oblog.setup(36,0)
	case 7
	enupload=oblog.setup(33,0)
	upfiletype=oblog.setup(34,0)
	onesize=oblog.setup(35,0)
	maxsize=oblog.setup(36,0)
	case 8
	enupload=oblog.setup(37,0)
	upfiletype=oblog.setup(38,0)
	onesize=oblog.setup(39,0)
	maxsize=oblog.setup(40,0)
	case 9
	enupload=oblog.setup(41,0)
	upfiletype=oblog.setup(42,0)
	onesize=oblog.setup(43,0)
	maxsize=oblog.setup(44,0)
end select
if enupload=0 then
	response.Write("��ǰϵͳ���ò������ϴ��ļ�")
	response.End()
end if
if  oblog.logined_uupfilemax>0 then	maxsize=oblog.logined_uupfilemax
freesize=int(maxsize-oblog.logined_uupfilesize/1024)
if freesize<=0 then
	response.Write("<ul style='margin:0px;text-align: left;width:100%;'> �ϴ��ռ��������������ϴ��ļ�,�������ϴ��ĵ�</ul></body></html>")
	response.End()
end if
If Request("t")="1" Then
	Upfile_Main()'1pho 0log;���ߴ����ύ����״̬ʱ��
Else
	Main()
End If

Sub Main()
	Dim PostRanNum
	Randomize
	PostRanNum = Int(900*rnd)+1000
	Session("UploadCode") = Cstr(PostRanNum)
%>
	<ul style="margin:0px;text-align: left;width:100%;"> 
      <form name="myform" method="post" action="upload.asp?t=1&tMode=<%=tMode%>&re=<%=re%>&ubb=<%=ubb%>&isphoto=<%=isphoto%>" enctype="multipart/form-data">
	<INPUT TYPE="hidden" NAME="UploadCode" value="<%=PostRanNum%>">	
	<input type="hidden" name="act" value="upload">
	�ϴ���Ƭ�� <input type="file" name="uploadfile">
	<input type="hidden" name="fname">
	<input type="submit" name="Ok" value="�ϴ�" > <!--ʣ��ռ䣺<%=freesize%>KB �����ļ���<%=onesize%>KB--> 
      </form>
	</ul>
</body>
</html>
<%
End Sub

Sub Upfile_Main()
%>
<ul style="margin:0px;text-align: left;width:100%;"> 
<%
UploadFile
%>
</ul>
</body>
</html>
<%
End Sub

Sub UploadFile()
	If Not oblog.ChkPost Then
		Exit Sub
	End If
	Server.ScriptTimeOut=9999999
'	'-----------------------------------------------------------------------------
	Dim Upload,FilePath,FormName,File,F_FileName,F_Viewname
	dim DrawInfo
	upfiletype=replace(upfiletype,"|",",")
	if freesize<=onesize then onesize=freesize
	if onesize<0 then onesize=0 
	if upload_dir<>"" then
		FilePath=upload_dir
	else
		FilePath = oblog.logined_udir&"/"&oblog.logined_ufolder&"/upload"
	end if
	FilePath=CreatePath(FilePath)
	If oblog.setup(63,0)="1" Then
		DrawInfo = oblog.setup(64,0)
	ElseIf oblog.setup(63,0)="2" Then
		DrawInfo = oblog.setup(69,0)
	Else
		DrawInfo = ""
	End If
	If DrawInfo = "0" Then
		DrawInfo = ""
		oblog.setup(63,0) = 0
	End If
	Set Upload = New UpFile_Cls
		Upload.UploadType			= Cint(oblog.setup(62,0))			'�����ϴ��������
		Upload.UploadPath			= FilePath								'�����ϴ�·��
		Upload.MaxSize				= Int(onesize)							'��λ KB
		Upload.InceptMaxFile		= 1										'ÿ���ϴ��ļ���������
		Upload.InceptFileType		= upfiletype							'�����ϴ��ļ�����
		Upload.RName				= "Homestay"&"-"
		Upload.ChkSessionName		= "UploadCode"
		
		if clng(oblog.setup(63,0))=1 or clng(oblog.setup(63,0))=2 then
			Upload.PreviewType			= 1								'����Ԥ��ͼƬ�������
		else
			Upload.PreviewType			= 999
		end if
		Upload.PreviewImageWidth	= 130										'����Ԥ��ͼƬ���
		Upload.PreviewImageHeight	= 100										'����Ԥ��ͼƬ�߶�
		Upload.DrawImageWidth		= oblog.setup(73,0)						'����ˮӡͼƬ������������
		Upload.DrawImageHeight		= oblog.setup(72,0)						'����ˮӡͼƬ����������߶�
		Upload.DrawGraph			= oblog.setup(70,0)						'����ˮӡ͸����
		Upload.DrawFontColor		= oblog.setup(66,0)						'����ˮӡ������ɫ
		Upload.DrawFontFamily		= oblog.setup(67,0)						'����ˮӡ���������ʽ
		Upload.DrawFontSize			= oblog.setup(65,0)						'����ˮӡ���������С
		Upload.DrawFontBold			= oblog.setup(68,0)						'����ˮӡ�����Ƿ����
		Upload.DrawInfo				=  DrawInfo								'����ˮӡ������Ϣ��ͼƬ��Ϣ
		Upload.DrawType				= oblog.setup(63,0)						'0=������ˮӡ ��1=����ˮӡ���֣�2=����ˮӡͼƬ
		Upload.DrawXYType			= oblog.setup(74,0)						'"0" =���ϣ�"1"=����,"2"=����,"3"=����,"4"=����
		Upload.DrawSizeType			= 1										'"0"=�̶���С��"1"=�ȱ�����С
		If oblog.setup(71,0)<>"" or oblog.setup(71,0)<>"0" Then
			Upload.TransitionColor	= oblog.setup(71,0)			'͸������ɫ����
		End If

		'ִ���ϴ�
		Upload.SaveUpFile
		If Upload.ErrCodes<>0 Then
			Response.write "����"& Upload.Description & "[ <a href='upload.asp?re="&re&"&ubb="&ubb&"&isphoto="&isphoto&"&tMode="& tMode &"'>�����ϴ�</a> ]"
			Exit Sub
		End If
		If Upload.Count > 0 Then
			For Each FormName In Upload.UploadFiles
				Set File = Upload.UploadFiles(FormName)
				F_FileName =  FilePath & File.FileName
				'����Ԥ����ˮӡͼƬ
				If Upload.PreviewType<>999 and File.FileType=1 then
						F_Viewname =  FilePath&"pre" & Replace(File.FileName,File.FileExt,"") & "jpg"
						'����Ԥ��ͼƬ:Call CreateView(ԭʼ�ļ���·��,Ԥ���ļ�����·��,ԭ�ļ���׺)
						Upload.CreateView F_FileName,F_Viewname,File.FileExt
				End If
					if re<>"no" then
						select case file.filetype
						case 1
							If tMode="1" Then
								Response.Write "<script>parent.oblogform.log_pics.value+='" & F_FileName & "\n';</script>"								
							End if
							if ubb="1" then
								Response.Write "<script>parent.oblogform.ubbedit.value+='[IMG]"&F_FileName&"[/IMG]';</script>"
							else
								Response.Write "<script>parent.oblog_InsertSymbol('<img src=" &F_FileName& "><br />');</script>"
							end if
							if isphoto=2 then
								Response.Write "<script>parent.oblogform.face.value='"&F_FileName&"';</script>"
							end if
						case else
							if ubb="1" then
								Response.Write "<script>parent.oblogform.ubbedit.value+='[UPLOAD="&blogdir&F_FileName&"]"&File.FileName&"[/UPLOAD]';</script>"
							else
								Response.Write "<script>parent.oblog_InsertSymbol('<a href=" &F_FileName& ">"&F_FileName&"</a>');</script>"
							end if
						end select
					else
						
					end if
					oblog.execute("update oblog_user set user_upfiles_size=user_upfiles_size+"&File.FileSize&" where userid="&oblog.logined_uid)
					oblog.execute("Insert into oblog_upfile (userid,file_name,file_path,file_ext,file_size,isphoto) values ("&oblog.logined_uid&",'"&File.FileName&"','"&F_FileName&"','"&File.FileExt&"',"&File.FileSize&","&isphoto&")")
					if instr("jpg,gif,bmp,pcx,png,psd",File.FileExt)=0 then isphoto=0
					if isphoto=1 then
						dim blog
						set blog=new class_blog
						blog.userid=oblog.logined_uid
						blog.update_album 0,0,0,""						
						'Response.Write "<script>alert(""�ϴ�ͼƬ�ɹ���"");parent.location.href=parent.location;<script>"
					else						
						'Response.Write "<script>alert(""�ϴ��ļ��ɹ���"");parent.location.href=parent.location;<script>"
					end if
					Response.Write "�ļ���"& F_FileName &"�ϴ��ɹ�![ <a href=upload.asp?re="&re&"&ubb="&ubb&"&isphoto="&isphoto&"&tMode=" & tMode &"><font color=#038ad7>�ٴ��ϴ�</font></a> ]"
				Set File = Nothing
			Next
		Else
			Response.write "����ȷѡ��Ҫ�ϴ����ļ���[ <a href='upload.asp?re="&re&"&ubb="&ubb&"&isphoto="&isphoto&"&tMode=" & tMode &"'><font color=#038ad7>�����ϴ�</font></a> ]"
			Exit Sub
		End If
	Set Upload = Nothing
End Sub

'����ϴ�Ŀ¼������Ŀ¼���Զ�����
Function CreatePath(PathValue)
	Dim objFSO,Fsofolder,uploadpath
	if upload_dir<>"" then
		uploadpath = "www.MyHomestay.com.cn"&"-"&year(Date) & "-" & month(Date)
	else
		uploadpath=""
	end if
	If Right(PathValue,1)<>"/" Then PathValue = PathValue&"/"
	On Error Resume Next
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		If objFSO.FolderExists(Server.MapPath(PathValue & uploadpath))=False Then
			objFSO.CreateFolder Server.MapPath(PathValue & uploadpath)
		End If
		If Err.Number = 0 and upload_dir<>"" Then
			CreatePath = PathValue & uploadpath & "/"
		Else
			CreatePath = PathValue
		End If
	Set objFSO = Nothing
End Function
%>