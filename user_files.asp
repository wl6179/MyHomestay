<!--#include file="user_top.asp"-->
<!--#include file="inc/class_blog.asp"-->


<div id="head">
	<div id="head2">
	
		<div id="head2-logo">
			
			<ul>
				<li class="active"><a href="http://www.myhomestay.com.cn">��������</a></li>
				<li><a href="http://www.myhomestay.com.cn">English</a></li>
				<!--<li><a href="http://www.myhomestay.com.cn">Janpan</a></li>-->
			</ul>
			
		</div>
		
		
		<div id="head2-menu">
			<div id="head2-search">
				<span>�ͻ���ѯ����</span>��<span>010-85493388</span>/<span>13910652850</span> &nbsp;&nbsp;&nbsp;
				<!--<a href="http://www.myhomestay.com.cn">��Ϊ��ҳ</a>&nbsp;&nbsp;&nbsp;-->
				<!--<a href="#" onClick="javascript:AddFavoriteOnClick();">���ո�������ղؼ�</a>-->&nbsp;
				<!--<a href="http://www.myhomestay.com.cn">��������</a>&nbsp;&nbsp;&nbsp;-->
			</div>
			
			<div id="divCN-EN">
			<ul>
				<li><a href="/index.asp">��ҳ<br />Home</a></li><!--#include file="menu/incIndexNav.htm"-->
	
				
			</ul>
			</div>
		</div>
	
	</div>
	
</div>
	
	
	<div id="body"><!--#body-->
	<div id="gbody"><!--#gbody-->
	
		<div id="left-margin"></div>
		
		<div id="main"><!--the main content area-->
			
			
			
			
			<div id="mainContentList">
				<div id="main-top-round">
					<div id="main-top-round-left"></div>
					<div id="main-top-round-right"></div>
				</div>
				
				<div id="main-main-round">
					<style>
						table {
						background:#ffffff;
						border:1px solid #999999;
						border-collapse: collapse;
						font-size:13px;
						}
						td {
						padding:0px 3px 0px 12px;
						width:100px;
						height:18px;
						border:1px solid #999999;
						}
						.tr01 {
						/*font:bold;*/
						background:#A1E5FA;
						font-boil:2px;
						color:#4f6b72;
						}
						.tr02 {
						background:#ffffff;
						color:#797268;
						}
						.tr03 {
						background:#F7FDFB;
						color:#4f6b72;
						}
					</style>
					<div id="main-TextArea01">
					
					
					
						<center><h2>�ļ���������</h2></center>
						
			


<%const MaxPerPage=24
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
strFileName="user_files.asp?usersearch=" & usersearch
if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if
select case action
	case "modifyphoto"
	call modify()
	case "savemodify"
	call Savemodify()
	case "delfile" 
	call delfile()
	case else
	call main()
end select
set rs=nothing
%>

</div>
  
  					<!--wl-->
					</div>
					
					
					<div id="main-main-round2">
						<center>
					  
					    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="585" height="90">
                          <param name="movie" value="images/90620_hp585_90.swf" />
                          <param name="quality" value="high" />
                          <embed src="images/90620_hp585_90.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="585" height="90"></embed>
					      </object>
					  	</center>
					</div>
					
					
		
					
					
				
					<div id="main-banner01">
						<div id="main-bottom-round-left"></div>
						<div id="main-bottom-round-right"></div>
					</div>
				</div>
				
				
				<div id="publicList">
					<div id="public-top-round">
						<div id="public-top-round-left"></div>
						<div id="public-top-round-right"></div>
					</div>
					
					<div id="public-main-round">
					<%'main()%>
						<div class="public-TextArea01">
							<p class="public-title01">
								��Ŀ��������
							</p>
							
							<!--#include file="menu/inc-menu.asp"-->
							
						</div>
						
					
					</div>
				
					<div id="public-bottom-round">
						<div id="public-bottom-round-left"></div>
						<div id="public-bottom-round-right"></div>
					</div>
				</div>				
			</div>
			

</div>
		</div><!--#gbody-->	
	  </div><!--#body-->


  

	  		
<%=oblog.user_copyright%>
			
			
	  	


  <script src="inc/divCN-EN.js" type="text/javascript"></script>
  </body>
</html>


<%
sub main()
	dim strGuide,ssql
	ssql="userid,file_name,file_path,file_size,fileid,file_readme,file_ext,isphoto,logid"
	strGuide="<h1 style='font-size:13px;'>��ǰ�ļ�����&nbsp;&gt;&gt;&nbsp;"
	select case usersearch
		case 0
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" order by fileid desc"
			end if
			strGuide=strGuide & "�����ļ�"
		case 1
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where file_ext='gif' or file_ext='jpg' or file_ext='bmp' or file_ext='png' or file_ext='psd' or file_ext='pcx' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='gif' or file_ext='jpg' or file_ext='bmp' or file_ext='png' or file_ext='psd' or file_ext='pcx' ) order by fileid desc"
			end if
			strGuide=strGuide & "ͼƬ�ļ�"
		case 2
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where file_ext='rar' or file_ext='zip' or file_ext='arj' or file_ext='sit' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='rar' or file_ext='zip' or file_ext='arj' or file_ext='sit') order by fileid desc"
			end if
			strGuide=strGuide & "ѹ���ļ�"
		case 3
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where file_ext='doc' or file_ext='xsl' or file_ext='txt' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='doc' or file_ext='xsl' or file_ext='txt') order by fileid desc"
			end if
			strGuide=strGuide & "�ĵ��ļ�"
		case 4
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where  file_ext='mp3' or file_ext='wmv' or file_ext='wma' or file_ext='rm' or file_ext='swf' order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and ( file_ext='mp3' or file_ext='wmv' or file_ext='wma' or file_ext='rm' or file_ext='swf') order by fileid desc"
			end if
			strGuide=strGuide & "ý���ļ�"
		case 5
			if oblog.logined_ulevel=9 then
				sql="select top 500 "&ssql&" from [oblog_upfile] where isphoto=1 order by fileid desc"
			else
				sql="select "&ssql&" from [oblog_upfile] where userid="&oblog.logined_uid&" and isphoto=1 order by fileid desc"
			end if
			strGuide=strGuide & "���"
		case else
	end select
	'strGuide=strGuide & "</td><td align='right'>"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	'response.Write(sql)
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		strGuide=strGuide & " (����0���ļ�)</h1><div id=""list"">"
		response.write strGuide
		showContent
	else
    	totalPut=rs.recordcount
		strGuide=strGuide & " (����" & totalPut & "���ļ�)</form></h1><div id=""list"">"
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
        	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���ļ�")
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark
            	showContent
            	response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���ļ�")
        	else
	        	currentPage=1
           		showContent
           		response.write oblog.showpage(strFileName,totalput,MaxPerPage,true,true,"���ļ�")
	    	end if
		end if
	end if
	rs.Close
	set rs=Nothing
end sub

sub showContent()
   	dim i,n,ext
    i=0
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
<style>	
	.list_right{
		width:100%;
	}
	.filepic {
		width:20%;
		float:left;
		margin:6px;
		
	}
	.fileimg {
		/*width:100%;*/
		
	}
</style>
   <div class="list_left">
   <form name="myform1" method="post" action="user_files.asp">
   <select size=1 name='usersearch' onChange='javascript:submit()'>
	<option value=10 selected>��ѡ���ļ�����</option> 
	<option value=0>�г������ļ�</option>
	<option value=1>ͼƬ�ļ�</option>
	<option value=2>ѹ���ļ�</option> 
	<option value=3>�ĵ��ļ�</option> 
	<option value=4>ý���ļ�</option>
	<option value=5>�����Ƭ</option>           
	</select>
	</form>
	<img align=absmiddle src='images/touming.gif' border="1" width='10' height='10' style='background:#cccccc;'>&nbsp;���ÿռ�:<%=oblog.showsize(oblog.logined_uupfilesize)%><br />
	<img align=absmiddle src='images/touming.gif' border="1" width='10' height='10' style='background:#ffffff;'>&nbsp;���ÿռ�:<%=oblog.showsize(freesize)%><br />
	<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="150" height="150">
	<param name=movie value="images/chart.swf?in6=<%=int((freesize/maxsize)*100)%>&in6=<%=int((oblog.logined_uupfilesize/maxsize)*100)%>">
	<param name=quality value=high>
	<param name="wmode" value="transparent">
	<embed src="images/chart.swf?in6=<%=int((freesize/maxsize)*100)%>&in6=<%=int(((maxsize-freesize)/maxsize)*100)%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="150" height="150">
	</embed> </object>
	</div>
    <div class="list_right"><iframe id=d_file frameborder=0 src="upload.asp?re=no" width="100%" height="30" scrolling=no></iframe><hr>
	<form name="myform" method="post" action="user_files.asp" onSubmit="return confirm('ȷ��Ҫɾ��ѡ�����ļ���');">
    <% do while not rs.eof 
		ext=rs("file_ext")%>
	<div class="filepic"><div class="fileimg"><a href='<%=rs("file_path")%>' target=_blank><%if rs("isphoto")=1 then response.Write("<img src=""images/filetype/photo.gif"" class=""fileimg"" alt=""����ļ�"" />") else showfilepic(rs("file_ext"))%></a></div>
	<span class="en"><%=oblog.showsize(rs("file_size"))%></span><br />
	<a href="user_files.asp?action=delfile&id=<%=rs("fileid")%>" onClick="return confirm('ȷ��Ҫɾ������ļ���');">ɾ��</a>
<%
'-----------------------------------
'ATFLAG
if  (ext="jpg" or ext="gif" or ext="pcx" or ext="bmp" or ext="png" or ext="psd") and en_photo=1 then
	if rs("isphoto")=0 then
		response.Write("<a href=""user_post.asp?t=1&photofile="&rs("file_path")&""">�������</a>")
	else
		response.Write("<a href=""user_post.asp?logid="&rs("logid")&"&t=1"">�޸�ͼƬ</a>")
	end if
end if
'-----------------------------------
%>
	<br /><input name='id' type='checkbox' onClick="unselectall()" id="id" value='<%=cstr(rs("fileid"))%>'><span class="en"><%="<a href="&rs("file_path")&" target=_blank>"&rs("file_name")&"</a>"%></span>
	</div>
<%  
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
	<ul class="list_bot" style="clear:both"> <input name="chkAll" type="checkbox" id="chkAll" onClick=CheckAll(this.form) value="checkbox" />ȫѡ
	<input name="action"  type="hidden" value="delfile" />
	<input type="submit" name="Submit" value="ɾ��ѡ�����ļ�" />
	</ul>
	</form>
  </div>  
 </div><br />
<%
end sub

function showfilepic(ext)
	ext=lcase(ext)
	if  ext="jpg" then
	response.Write("<img src=""images/filetype/jpg.gif"" class=""fileimg"" alt=""JPG�ļ�"" />")
	elseif  ext="gif" then
	response.Write("<img src=""images/filetype/gif.gif"" class=""fileimg"" alt=""GIF�ļ�"" />")
	elseif  ext="bmp" then
	response.Write("<img src=""images/filetype/bmp.gif"" class=""fileimg"" alt=""BMP�ļ�"" />")
	elseif  ext="png" then
	response.Write("<img src=""images/filetype/png.gif"" class=""fileimg"" alt=""PNG�ļ�"" />")
	elseif  ext="psd" then
	response.Write("<img src=""images/filetype/psd.gif"" class=""fileimg"" alt=""PSD�ļ�"" />")
	elseif ext="rar" or ext="zip" or ext="arj" or ext="sit" then
	response.Write("<img src=""images/filetype/rar.gif"" class=""fileimg"" alt=""ѹ���ļ�"" />")
	elseif ext="xsl" then
	response.Write("<img src=""images/filetype/excel.gif"" class=""fileimg"" alt=""Excel�ļ�"" />")
	elseif ext="doc" then
	response.Write("<img src=""images/filetype/word.gif"" class=""fileimg"" alt=""Word�ļ�"" />")
	elseif ext="mp3" then
	response.Write("<img src=""images/filetype/mp3.gif"" class=""fileimg"" alt=""mp3�ļ�"" />")
	elseif ext="rm" or ext="ram" then
	response.Write("<img src=""images/filetype/rm.gif"" class=""fileimg"" alt=""Real�ļ�"" />")
	elseif ext="wmv" or ext="wma" or ext="mpg" or ext="avi" then
	response.Write("<img src=""images/filetype/media.gif"" class=""fileimg"" alt=""ý���ļ�"" />")
	else
	response.Write("<img src=""images/filetype/blank.gif"" class=""fileimg"" alt=""�����ļ�"" />")
	end if
end function

sub delfile()
	if id="" then
		oblog.adderrstr( "������ָ��Ҫɾ�����ļ���")
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
	oblog.showok "ɾ���ļ��ɹ���",""
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
	end if
end sub
%>