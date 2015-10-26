<!--#include file="inc/inc_sys.asp"-->
<%
dim Action,FoundErr,ErrMsg
Action=trim(request("Action"))
dim dbpath
dim ObjInstalled
if not IsObject(conn) then link_database
if is_sqldata=0 then dbpath=server.mappath(db)
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")

%>
<html>
<head>
<title>数据库备份</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="images/admin/Admin_STYLE.CSS">
<style type="text/css">
<!--
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
  <tr class="topbg"> 
    <td height="22" colspan="2" align="center"><strong>数 据 库 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="admin_database.asp?Action=Backup">备份数据库</a> | <a href="admin_database.asp?Action=Restore">恢复数据库</a> 
      | <a href="admin_database.asp?Action=Compact">压缩数据库</a> | <a href="admin_database.asp?Action=SpaceSize">系统空间占用情况</a></td>
  </tr>
</table>
<%
if Action="Backup" or Action="BackupData" then
	if isobject(conn) then conn.close:set conn=nothing
	call ShowBackup()
elseif Action="Compact" or Action="CompactData" then
	if isobject(conn) then conn.close:set conn=nothing
	call ShowCompact()
elseif Action="Restore" or Action="RestoreData" then
	if isobject(conn) then conn.close:set conn=nothing
	call ShowRestore()
elseif Action="SpaceSize" then
	call SpaceSize()
	if isobject(conn) then conn.close:set conn=nothing
else
	FoundErr=True
	ErrMsg=ErrMsg & "<br><li>错误参数！</li>"
	if isobject(conn) then conn.close:set conn=nothing
end if
if FoundErr=True then
	call WriteErrMsg()
end if

sub ShowBackup()
if is_sqldata=1 then sqldata_readme
%>
<form method="post" action="admin_database.asp?action=BackupData">
<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0" class="border">
  <tr class="title"> 
      <td align="center" height="22" valign="middle"><span class="style1">备 份 数 据 库</span></td>
  </tr>
  <tr class="tdbg"> 
    <td height="150" align="center" valign="middle">
<%    
if request("action")="BackupData" then
	call backupdata()
else
%>
        <table cellpadding="3" cellspacing="1" border="0" width="100%">
          <tr> 
            <td width="200" height="33" align="right">备份目录：</td>
            <td><input type=text size=20 name=bkfolder value=Databackup></td>
            <td>相对路径目录，如目录不存在，将自动创建</td>
          </tr>
          <tr> 
            <td width="200" height="34" align="right">备份名称：</td>
            <td height="34"><input type=text size=20 name=bkDBname value="<%=date()%>"></td>
            <td height="34">不用输入文件名后缀（默认为“.asa”）。如有同名文件，将覆盖</td>
          </tr>
          <tr align="center"> 
            <td height="40" colspan="3"><input name="submit" type=submit value=" 开始备份 " <%If ObjInstalled=false Then response.Write "disabled"%>></td>
          </tr>
        </table>
<%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
	end if
end if
%>
    </td>
 </tr>
</table>
</form>
<%
end sub

sub ShowCompact()
if is_sqldata=1 then sqldata_readme
%>
<form method="post" action="admin_database.asp?action=CompactData">
<table class="border" width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr class="title"> 
    <td align="center" height="22" valign="middle"><span class="style1">数据库在线压缩</span></td>
  </tr>
  <tr class="tdbg"> 
    <td align="center" height="150" valign="middle"> 
      <%    
if request("action")="CompactData" then
	call CompactData()
else
%>
      <br> <br> <br>
      压缩前，建议先备份数据库，以免发生意外错误。 <br> <br> <br> <input name="submit2" type=submit value=" 压缩数据库 " <%If ObjInstalled=false Then response.Write "disabled"%>> <br> <br> 
      <%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
	end if
end if
%>
    </td>
  </tr>
</table>
</form>
<%
end sub

sub ShowRestore()
if is_sqldata=1 then sqldata_readme
%>
<form method="post" action="admin_database.asp?action=RestoreData">
	<table width="98%" class="border" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr class="title"> 
      		<td align="center" height="22" valign="middle"><span class="style1">数据库恢复</span></td>
        </tr>
        <tr class="tdbg">
            <td align="center" height="150" valign="middle"> 
        <%
if request("action")="RestoreData" then
	call RestoreData()
else
%>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="200" height="30" align="right">备份数据库路径（相对）：</td>
            <td height="30"><input name=backpath type=text id="backpath" value="DataBackup\oblog.mdb" size=50 maxlength="200"></td>
          </tr>
          <tr align="center"> 
            <td height="40" colspan="2"><input name="submit" type=submit value=" 恢复数据 " <%If ObjInstalled=false Then response.Write "disabled"%>></td>
          </tr>
        </table>
<%
	If ObjInstalled=false Then
		Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)! 不能使用本功能</font></b>"
	end if
end if
%>
            </td>
        </tr>
  </table>
</form>
<%
end sub


sub SpaceSize()
	on error resume next
	dim rs
	set rs=oblog.execute("select userdir from oblog_userdir")
%>
<br>
<table class="border" width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
	<tr class="title"> 
		<td align="center" height="22" valign="middle"><span class="style1">系统空间占用情况</span></td>
	</tr>
  <tr class="tdbg"> 
    <td width="100%" height="150" valign="middle">
	<blockquote><br>
        系统数据占用空间：&nbsp;&nbsp; 
        <%showSpaceinfo("data")%>
      <br>
      <br>
        备份数据占用空间：&nbsp;&nbsp; 
        <%showSpaceinfo("databackup")%>
      <br>
      <br>
        程序文件占用空间：&nbsp;&nbsp; 
        <%showSpecialSpaceinfo("Program")%>
      <br>
      <br>
        配色模板占用空间：&nbsp;&nbsp; 
        <%showSpaceinfo("skin")%>
      <br>
      <br>
        系统图片占用空间：&nbsp;&nbsp; 
        <%showSpaceinfo("images")%>
	  <%while not rs.eof %>
      <br>
      <br>
      <%=rs(0)%>目录占用空间：&nbsp;&nbsp; 
        <%showSpaceinfo(rs(0))%>
	  <%rs.movenext
	  wend%>
      <br>
      <br>
        系统占用空间总计： 
        <%showspecialspaceinfo("All")%>
	</blockquote> 	
    </td>
  </tr>
</table>
<br>
<%
end sub
sub sqldata_readme
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="98%" class="border">			<tr>
  					
    <th height=25> &nbsp;<font color="#FFFFFF">&nbsp;SQL数据库数据处理说明</font> </th>
  				</tr> 	
 				<tr class="tdbg">
 					<td > 			
 			<blockquote>
<B>一、备份数据库</B>
<BR><BR>
1、打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server<BR>
2、SQL Server组-->双击打开你的服务器-->双击打开数据库目录<BR>
3、选择你的数据库名称（如论坛数据库Forum）-->然后点上面菜单中的工具-->选择备份数据库<BR>
4、备份选项选择完全备份，目的中的备份到如果原来有路径和名称则选中名称点删除，然后点添加，如果原来没有路径和名称则直接选择添加，接着指定路径和文件名，指定后点确定返回备份窗口，接着点确定进行备份
<BR><BR>
<B>二、还原数据库</B><BR><BR>
1、打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server<BR>
2、SQL Server组-->双击打开你的服务器-->点图标栏的新建数据库图标，新建数据库的名字自行取<BR>
3、点击新建好的数据库名称（如论坛数据库Forum）-->然后点上面菜单中的工具-->选择恢复数据库<BR>
4、在弹出来的窗口中的还原选项中选择从设备-->点选择设备-->点添加-->然后选择你的备份文件名-->添加后点确定返回，这时候设备栏应该出现您刚才选择的数据库备份文件名，备份号默认为1（如果您对同一个文件做过多次备份，可以点击备份号旁边的查看内容，在复选框中选择最新的一次备份后点确定）-->然后点击上方常规旁边的选项按钮<BR>
5、在出现的窗口中选择在现有数据库上强制还原，以及在恢复完成状态中选择使数据库可以继续运行但无法还原其它事务文章的选项。在窗口的中间部位的将数据库文件还原为这里要按照你SQL的安装进行设置（也可以指定自己的目录），逻辑文件名不需要改动，移至物理文件名要根据你所恢复的机器情况做改动，如您的SQL数据库装在D:\Program Files\Microsoft SQL Server\MSSQL\Data，那么就按照您恢复机器的目录进行相关改动改动，并且最后的文件名最好改成您当前的数据库名（如原来是bbs_data.mdf，现在的数据库是forum，就改成forum_data.mdf），文章和数据文件都要按照这样的方式做相关的改动（文章的文件名是*_log.ldf结尾的），这里的恢复目录您可以自由设置，前提是该目录必须存在（如您可以指定d:\sqldata\bbs_data.mdf或者d:\sqldata\bbs_log.ldf），否则恢复将报错<BR>
6、修改完成后，点击下面的确定进行恢复，这时会出现一个进度条，提示恢复的进度，恢复完成后系统会自动提示成功，如中间提示报错，请记录下相关的错误内容并询问对SQL操作比较熟悉的人员，一般的错误无非是目录错误或者文件名重复或者文件名错误或者空间不够或者数据库正在使用中的错误，数据库正在使用的错误您可以尝试关闭所有关于SQL窗口然后重新打开进行恢复操作，如果还提示正在使用的错误可以将SQL服务停止然后重起看看，至于上述其它的错误一般都能按照错误内容做相应改动后即可恢复<BR><BR>

<B>三、收缩数据库</B><BR><BR>
一般情况下，SQL数据库的收缩并不能很大程度上减小数据库大小，其主要作用是收缩文章大小，应当定期进行此操作以免数据库文章过大<BR>
1、设置数据库模式为简单模式：打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器-->双击打开数据库目录-->选择你的数据库名称（如论坛数据库Forum）-->然后点击右键选择属性-->选择选项-->在故障还原的模式中选择“简单”，然后按确定保存<BR>
2、在当前数据库上点右键，看所有任务中的收缩数据库，一般里面的默认设置不用调整，直接点确定<BR>
3、<font color=blue>收缩数据库完成后，建议将您的数据库属性重新设置为标准模式，操作方法同第一点，因为文章在一些异常情况下往往是恢复数据库的重要依据</font>
<BR><BR>

<B>四、设定每日自动备份数据库</B><BR><BR>
<font color=red>强烈建议有条件的用户进行此操作！</font><BR>
1、打开企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器<BR>
2、然后点上面菜单中的工具-->选择数据库维护计划器<BR>
3、下一步选择要进行自动备份的数据-->下一步更新数据优化信息，这里一般不用做选择-->下一步检查数据完整性，也一般不选择<BR>
4、下一步指定数据库维护计划，默认的是1周备份一次，点击更改选择每天备份后点确定<BR>
5、下一步指定备份的磁盘目录，选择指定目录，如您可以在D盘新建一个目录如：d:\databak，然后在这里选择使用此目录，如果您的数据库比较多最好选择为每个数据库建立子目录，然后选择删除早于多少天前的备份，一般设定4－7天，这看您的具体备份要求，备份文件扩展名一般都是bak就用默认的<BR>
6、下一步指定事务文章备份计划，看您的需要做选择-->下一步要生成的报表，一般不做选择-->下一步维护计划历史记录，最好用默认的选项-->下一步完成<BR>
7、完成后系统很可能会提示Sql Server Agent服务未启动，先点确定完成计划设定，然后找到桌面最右边状态栏中的SQL绿色图标，双击点开，在服务中选择Sql Server Agent，然后点击运行箭头，选上下方的当启动OS时自动启动服务<BR>
8、这个时候数据库计划已经成功的运行了，他将按照您上面的设置进行自动备份
<BR><BR>
修改计划：<BR>
1、打开企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器-->管理-->数据库维护计划-->打开后可看到你设定的计划，可以进行修改或者删除操作
<BR><BR>
<B>五、数据的转移（新建数据库或转移服务器）</B><BR><BR>
一般情况下，最好使用备份和还原操作来进行转移数据，在特殊情况下，可以用导入导出的方式进行转移，这里介绍的就是导入导出方式，导入导出方式转移数据一个作用就是可以在收缩数据库无效的情况下用来减小（收缩）数据库的大小，本操作默认为您对SQL的操作有一定的了解，如果对其中的部分操作不理解，可以咨询动网相关人员或者查询网上资料<BR>
1、将原数据库的所有表、存储过程导出成一个SQL文件，导出的时候注意在选项中选择编写索引脚本和编写主键、外键、默认值和检查约束脚本选项<BR>
2、新建数据库，对新建数据库执行第一步中所建立的SQL文件<BR>
3、用SQL的导入导出方式，对新数据库导入原数据库中的所有表内容<BR>
 			</blockquote> 	
 					</td>
 				</tr>
 			</table>
<%end sub%>
</body>
</html>
<%
sub BackupData()
	dim bkfolder,bkdbname,fso
	bkfolder=trim(request("bkfolder"))
	bkdbname=trim(request("bkdbname"))
	if bkfolder="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定备份目录！</li>"
	end if
	if bkdbname="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定备份文件名</li>"
	end if
	if FoundErr=True then exit sub
	bkfolder=server.MapPath(bkfolder)
	Set Fso=server.createobject("scripting.filesystemobject")
	if fso.FileExists(dbpath) then
		If fso.FolderExists(bkfolder)=false Then
			fso.CreateFolder(bkfolder)
		end if
		fso.copyfile dbpath,bkfolder & "\" & bkdbname & ".asa"
		response.write "<center>备份数据库成功，备份的数据库为 " & bkfolder & "\" & bkdbname & ".asa</center>"
	Else
		response.write "<center>找不到源数据库文件，请检查inc/conn.asp中的配置。</center>"
	End if
end sub

sub CompactData()
	Dim fso, Engine, strDBPath
	strDBPath = left(dbPath,instrrev(DBPath,"\"))
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(dbPath) Then
		Set Engine = CreateObject("JRO.JetEngine")
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath," Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
		fso.CopyFile strDBPath & "temp.mdb",dbpath
		fso.DeleteFile(strDBPath & "temp.mdb")
		Set fso = nothing
		Set Engine = nothing
		response.write "数据库压缩成功!"
	Else
		response.write "数据库没有找到!"
	End If
end sub

sub RestoreData()
	dim backpath,fso
	backpath=request.form("backpath")
	if backpath="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定原备份的数据库文件名！<li>"
		exit sub	
	end if
	backpath=server.mappath(backpath)
	Set Fso=server.createobject("scripting.filesystemobject")
	If Right(LCase(backpath),4)=".asa" OR Right(LCase(backpath),4)=".asp" Or Right(LCase(backpath),4)=".mdb" Then		
		if fso.fileexists(backpath) then  					
			fso.copyfile Backpath,Dbpath
			response.write "成功恢复数据！"
		else
			response.write "找不到指定的备份文件！"	
		end if
	Else
		response.write "错误的备份文件！"	
	End If
	Set Fso=Nothing
end sub

Sub ShowSpaceInfo(drvpath)
	dim fso,d,size,showsize
	set fso=server.createobject("scripting.filesystemobject") 		
	drvpath=server.mappath(drvpath) 		 		
	set d=fso.getfolder(drvpath) 		
	size=d.size
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(size\1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
End Sub	

Sub Showspecialspaceinfo(method)
	dim fso,d,fc,f1,size,showsize,drvpath 		
	set fso=server.createobject("scripting.filesystemobject")
	drvpath=server.mappath("pic")
	drvpath=left(drvpath,(instrrev(drvpath,"\")-1))
	set d=fso.getfolder(drvpath) 		
	
	if method="All" then 		
		size=d.size
	elseif method="Program" then
		set fc=d.Files
		for each f1 in fc
			size=size+f1.size
		next	
	end if	
	
	showsize=size & "&nbsp;Byte" 
	if size>1024 then
	   size=(size\1024)
	   showsize=size & "&nbsp;KB"
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;MB"		
	end if
	if size>1024 then
	   size=(size/1024)
	   showsize=formatnumber(size,2) & "&nbsp;GB"	   
	end if   
	response.write "<font face=verdana>" & showsize & "</font>"
end sub 	 	 	

Function Drawbar(drvpath)
	dim fso,drvpathroot,d,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("pic")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	
	drvpath=server.mappath(drvpath) 		
	set d=fso.getfolder(drvpath)
	size=d.size
	
	barsize=cint((size/totalsize)*400)
	Drawbar=barsize
End Function 	

Function Drawspecialbar()
	dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
	set fso=server.createobject("scripting.filesystemobject")
	drvpathroot=server.mappath("pic")
	drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
	set d=fso.getfolder(drvpathroot)
	totalsize=d.size
	
	set fc=d.files
	for each f1 in fc
		size=size+f1.size
	next	
	
	barsize=cint((size/totalsize)*400)
	Drawspecialbar=barsize
End Function 	

'**************************************************
'函数名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'       False ----没有安装
'**************************************************
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>