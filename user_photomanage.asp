<!--#include file="user_top.asp"-->
<!--#include file="inc/class_blog.asp"-->
<div id="main">
  <div class="submenu">
  	<div class="side_c1 side11"></div>
    <div class="side_c2 side21"></div>
	<div class="submenu_content">
		&#8226; <a href='user_post.asp?t=1'>添加图片</a>
		&#8226; <a href='user_photomanage.asp'>相册图片管理</a>
		&#8226; <a href='user_blogmanage.asp?t=1'>相册内容管理</a>
		&#8226;<a href="user_subject.asp?t=1">相册分类管理</a>
		
	</div>
  </div>
  <div class="content">
  	<div class="content_top">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
			相册管理
	</div>
	
    <div class="content_body">
<%
if en_photo=0 then
	oblog.adderrstr("此功能已被系统关闭！")
	oblog.showusererr
end if
const MaxPerPage=24
dim strFileName,totalPut,TotalPages
dim rs, sql,action
dim id,usersearch,Keyword,strField
keyword=trim(request("keyword"))
if keyword<>"" then 
	keyword=oblog.filt_badstr(keyword)
end if
strField=trim(request("Field"))
usersearch=trim(request("usersearch"))
action=trim(request("action"))
id=oblog.filt_badstr(trim(Request("id")))
if usersearch="" then
	usersearch=0
else
	usersearch=Clng(usersearch)
end if
strFileName="user_photomanage.asp?usersearch=" & usersearch
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
select case action
	case "delfile" 
		call delfile()
	case else
	call main()
end select
set rs=nothing
%>
	</div>
	
    <div class="content_bot">
		  	<div class="side_e1 side12"></div>
   			<div class="side_e2 side22"></div>
 	</div>
	
  </div>
</div>  
  
  
  <div id="bottom"><%=oblog.user_copyright%></div>

</body>
</html>
<%
sub main()
	dim strGuide,ssql
	ssql="userid,file_name,file_path,file_size,fileid,file_readme,file_ext,isphoto,logid"
	strGuide="<h1>我的相册"
	if oblog.logined_ulevel=9 then
		sql="select top 500 "&ssql&" from [oblog_upfile] where isphoto=1 order by fileid desc"
	else
		sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and isphoto=1 order by fileid desc"
	end if
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'response.Write(sql)
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & " (共有0个文件)</h1><div id=""list"">"
		response.write strGuide
		showContent
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & " (共有" & totalPut & "个文件)</form></h1><div id=""list"">"
		response.write strGuide
		if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if
    	end if
	    if currentPage=1 then
        	showContent
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个文件")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个文件")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"个文件")
	    	end if
		end if
	end if
	rs.Close
	set rs=Nothing
end sub

sub showContent()
   	dim i,n,ext,imgsrc,fso
    i=0
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	dim freesize,maxsize
	select case oblog.logined_ulevel
		case 7
		maxsize=oblog.setup(36,0)
		case 8
		maxsize=oblog.setup(40,0)
		case 9
		maxsize=oblog.setup(44,0)
	end select
	if  oblog.logined_uupfilemax>0 then	maxsize=oblog.logined_uupfilemax
	maxsize=maxsize*1024
	freesize=int(maxsize-oblog.logined_uupfilesize)

%>
	
   <div class="list_left">
	<img align=absmiddle src='images/touming.gif' border="1" width='10' height='10' style='background:#cccccc;'>&nbsp;已用空间:<%=oblog.showsize(oblog.logined_uupfilesize)%><br />
	<img align=absmiddle src='images/touming.gif' border="1" width='10' height='10' style='background:#ffffff;'>&nbsp;可用空间:<%=oblog.showsize(freesize)%><br />
	<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="150" height="150">
	<param name=movie value="images/chart.swf?in6=<%=int((freesize/maxsize)*100)%>&in6=<%=int((oblog.logined_uupfilesize/maxsize)*100)%>">
	<param name=quality value=high>
	<param name="wmode" value="transparent">
	<embed src="images/chart.swf?in6=<%=int((freesize/maxsize)*100)%>&in6=<%=int(((maxsize-freesize)/maxsize)*100)%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="150" height="150">
	</embed> </object>
	</div>
	<div class="list_right">
	<%
	If t="1" Then
	%>
    <iframe id=d_file frameborder=0 src="upload.asp?re=no&isphoto=1" width="100%" height="30" scrolling=no></iframe><hr>
    <%
	End If
    %>
	<form name="myform" method="post" action="user_photomanage.asp" onSubmit="return confirm('确定要删除选定的文件吗？');">
    <% do while not rs.eof 
		ext=rs("file_ext")
		imgsrc=rs("file_path")
		imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
		imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
		if  not fso.FileExists(Server.MapPath(imgsrc)) then
			imgsrc=rs("file_path")
		end if%>
	<div class="photopic"><div class="photoimg_div"><a href='<%=rs("file_path")%>' target=_blank><%if rs("isphoto")=1 then response.Write("<img src="""&imgsrc&""" class=""photoimg"" alt=""相册文件"" Width=180 height=120 />") %></a>
	<span class="en"><%=oblog.showsize(rs("file_size"))%>
	<a href="user_photomanage.asp?action=delfile&id=<%=rs("fileid")%>" onClick="return confirm('确定要删除这个文件吗？');">删除</a>
	</span>
<%
if  ext="jpg" or ext="gif" or ext="pcx" or ext="bmp" or ext="png" or ext="psd" then
	if rs("isphoto")=0 then
		response.Write("<a href=""javascript:openScript('user_photo.asp?action=modifyphoto&id="&rs("fileid")&"',550,600)"">加入相册</a>")
	else
		response.Write("<a href=""user_post.asp?logid="&rs("logid")&"&t=1"">修改图片</a>")
	end if
end if
%>
	</div><br /><input name='id' type='checkbox' onClick="unselectall()" id="id" value='<%=cstr(rs("fileid"))%>'><span class="en"><%="<a href="&rs("file_path")&" target=_blank>"&rs("file_name")&"</a>"%></span>
	</div>
<%  
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
	<ul class="list_bot"> <input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" />全选
	<input name="action"  type="hidden" value="delfile" />
	<input type="submit" name="Submit" value="删除选定的文件" />
	</ul>
	</form>
  </div>  
 </div>
<%
	set fso=nothing
end sub

function showfilepic(ext)
	ext=lcase(ext)
	if  ext="jpg" then
	response.Write("<img src=""images/filetype/jpg.gif"" class=""fileimg"" alt=""JPG文件"" />")
	elseif  ext="gif" then
	response.Write("<img src=""images/filetype/gif.gif"" class=""fileimg"" alt=""GIF文件"" />")
	elseif  ext="bmp" then
	response.Write("<img src=""images/filetype/bmp.gif"" class=""fileimg"" alt=""BMP文件"" />")
	elseif  ext="png" then
	response.Write("<img src=""images/filetype/png.gif"" class=""fileimg"" alt=""PNG文件"" />")
	elseif  ext="psd" then
	response.Write("<img src=""images/filetype/psd.gif"" class=""fileimg"" alt=""PSD文件"" />")
	elseif ext="rar" or ext="zip" or ext="arj" or ext="sit" then
	response.Write("<img src=""images/filetype/rar.gif"" class=""fileimg"" alt=""压缩文件"" />")
	elseif ext="xsl" then
	response.Write("<img src=""images/filetype/excel.gif"" class=""fileimg"" alt=""Excel文件"" />")
	elseif ext="doc" then
	response.Write("<img src=""images/filetype/word.gif"" class=""fileimg"" alt=""Word文件"" />")
	elseif ext="mp3" then
	response.Write("<img src=""images/filetype/mp3.gif"" class=""fileimg"" alt=""mp3文件"" />")
	elseif ext="rm" or ext="ram" then
	response.Write("<img src=""images/filetype/rm.gif"" class=""fileimg"" alt=""Real文件"" />")
	elseif ext="wmv" or ext="wma" or ext="mpg" or ext="avi" then
	response.Write("<img src=""images/filetype/media.gif"" class=""fileimg"" alt=""媒体文件"" />")
	else
	response.Write("<img src=""images/filetype/blank.gif"" class=""fileimg"" alt=""其他文件"" />")
	end if
end function

sub delfile()
	if id="" then
		oblog.adderrstr( "错误：请指定要删除的文件！")
		oblog.showusererr
		exit sub
	end if
	if instr(id,",")>0 then
		dim n,i
		id=FilterIDs(id)
		n=split(id,",")
		for i=0 to ubound(n)
			delonefile(n(i))
		next
	else
		delonefile(id)
	end if
	set rs=nothing
	oblog.showok "删除文件成功！",""
end sub

sub delonefile(id)
	id=clng(id)
	dim userid,filesize,filepath,fso,photofile,isphoto,imgsrc
	if oblog.logined_ulevel=9 then
		sql="select * from [oblog_upfile] where fileid=" & id
	else
		sql="select * from [oblog_upfile] where fileid=" & id&" and userid="&oblog.logined_uid
	end if
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,3
	if not rs.eof then
		userid=rs("userid")
		filesize=clng(rs("file_size"))
		filepath=rs("file_path")
		photofile=rs("photofile")
		isphoto=rs("isphoto")
		rs.delete
		rs.update
		rs.close
		oblog.execute("update [oblog_user] set user_upfiles_size=user_upfiles_size-"&filesize&" where userid="&userid)
		if filepath<>"" then
			imgsrc=filepath
			Set fso = server.CreateObject("Scripting.FileSystemObject")
			if instr("jpg,bmp,gif,png,pcx",right(imgsrc,3))>0 then
				imgsrc=replace(imgsrc,right(imgsrc,3),"jpg")
				imgsrc=replace(imgsrc,right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")),"pre"&right(imgsrc,len(imgsrc)-InstrRev(imgsrc,"/")))
				if  fso.FileExists(Server.MapPath(imgsrc)) then
					fso.DeleteFile Server.MapPath(imgsrc)
				end if
			end if
			if fso.FileExists(Server.MapPath(filepath)) then 
				fso.DeleteFile Server.MapPath(filepath)
			end if
			set fso=nothing
		end if
	else
		rs.close
		set rs=nothing
	End If
end sub

%>