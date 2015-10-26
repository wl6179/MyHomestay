<!-- #include file="inc/inc_syssite.asp" -->
<!-- #include File="inc/class_upfile.asp" -->
<!-- #include File="inc/class_blog.asp" -->
	<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
	<title>MyHomestay 文件上传</title>
	<link href="css/back-ctrl.css" rel="stylesheet" type="text/css" />
	</head>
	<body>

<%
if not oblog.checkuserlogined() then
	response.Write("登陆后才能上传文件")
	response.End()
end if
Dim tMode
tMode=Request("tMode")'0log 1pho;
dim freesize,onesize,maxsize,enupload,upfiletype,re,ubb,isphoto
re=request.QueryString("re")
ubb=request.QueryString("ubb")
isphoto=cint(request.QueryString("isphoto"))
select case oblog.logined_ulevel
	case 6'WL中国家庭上传参数；
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
	response.Write("当前系统设置不允许上传文件")
	response.End()
end if
if  oblog.logined_uupfilemax>0 then	maxsize=oblog.logined_uupfilemax
freesize=int(maxsize-oblog.logined_uupfilesize/1024)
if freesize<=0 then
	response.Write("<ul style='margin:0px;text-align: left;width:100%;'> 上传空间已满，不允许上传文件,请整理上传文档</ul></body></html>")
	response.End()
end if
If Request("t")="1" Then
	Upfile_Main()'1pho 0log;或者处于提交处理状态时；
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
	上传照片： <input type="file" name="uploadfile">
	<input type="hidden" name="fname">
	<input type="submit" name="Ok" value="上传" > <!--剩余空间：<%=freesize%>KB 单个文件：<%=onesize%>KB--> 
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
		Upload.UploadType			= Cint(oblog.setup(62,0))			'设置上传组件类型
		Upload.UploadPath			= FilePath								'设置上传路径
		Upload.MaxSize				= Int(onesize)							'单位 KB
		Upload.InceptMaxFile		= 1										'每次上传文件个数上限
		Upload.InceptFileType		= upfiletype							'设置上传文件限制
		Upload.RName				= "Homestay"&"-"
		Upload.ChkSessionName		= "UploadCode"
		
		if clng(oblog.setup(63,0))=1 or clng(oblog.setup(63,0))=2 then
			Upload.PreviewType			= 1								'设置预览图片组件类型
		else
			Upload.PreviewType			= 999
		end if
		Upload.PreviewImageWidth	= 130										'设置预览图片宽度
		Upload.PreviewImageHeight	= 100										'设置预览图片高度
		Upload.DrawImageWidth		= oblog.setup(73,0)						'设置水印图片或文字区域宽度
		Upload.DrawImageHeight		= oblog.setup(72,0)						'设置水印图片或文字区域高度
		Upload.DrawGraph			= oblog.setup(70,0)						'设置水印透明度
		Upload.DrawFontColor		= oblog.setup(66,0)						'设置水印文字颜色
		Upload.DrawFontFamily		= oblog.setup(67,0)						'设置水印文字字体格式
		Upload.DrawFontSize			= oblog.setup(65,0)						'设置水印文字字体大小
		Upload.DrawFontBold			= oblog.setup(68,0)						'设置水印文字是否粗体
		Upload.DrawInfo				=  DrawInfo								'设置水印文字信息或图片信息
		Upload.DrawType				= oblog.setup(63,0)						'0=不加载水印 ，1=加载水印文字，2=加载水印图片
		Upload.DrawXYType			= oblog.setup(74,0)						'"0" =左上，"1"=左下,"2"=居中,"3"=右上,"4"=右下
		Upload.DrawSizeType			= 1										'"0"=固定缩小，"1"=等比例缩小
		If oblog.setup(71,0)<>"" or oblog.setup(71,0)<>"0" Then
			Upload.TransitionColor	= oblog.setup(71,0)			'透明度颜色设置
		End If

		'执行上传
		Upload.SaveUpFile
		If Upload.ErrCodes<>0 Then
			Response.write "错误："& Upload.Description & "[ <a href='upload.asp?re="&re&"&ubb="&ubb&"&isphoto="&isphoto&"&tMode="& tMode &"'>重新上传</a> ]"
			Exit Sub
		End If
		If Upload.Count > 0 Then
			For Each FormName In Upload.UploadFiles
				Set File = Upload.UploadFiles(FormName)
				F_FileName =  FilePath & File.FileName
				'创建预览及水印图片
				If Upload.PreviewType<>999 and File.FileType=1 then
						F_Viewname =  FilePath&"pre" & Replace(File.FileName,File.FileExt,"") & "jpg"
						'创建预览图片:Call CreateView(原始文件的路径,预览文件名及路径,原文件后缀)
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
						'Response.Write "<script>alert(""上传图片成功！"");parent.location.href=parent.location;<script>"
					else						
						'Response.Write "<script>alert(""上传文件成功！"");parent.location.href=parent.location;<script>"
					end if
					Response.Write "文件："& F_FileName &"上传成功![ <a href=upload.asp?re="&re&"&ubb="&ubb&"&isphoto="&isphoto&"&tMode=" & tMode &"><font color=#038ad7>再次上传</font></a> ]"
				Set File = Nothing
			Next
		Else
			Response.write "请正确选择要上传的文件。[ <a href='upload.asp?re="&re&"&ubb="&ubb&"&isphoto="&isphoto&"&tMode=" & tMode &"'><font color=#038ad7>重新上传</font></a> ]"
			Exit Sub
		End If
	Set Upload = Nothing
End Sub

'检查上传目录，若无目录则自动建立
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