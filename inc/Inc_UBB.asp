<%
	dim CurrentPage '首页显示分页使用
	'根据分类id，返回分类名
	Public Function getsubname(str,id)
		'ON ERROR RESUME NEXT
        Dim tmp1, tmp2
        id = Trim(id)
        tmp1 = InStr(Str, id & "!!??((")
        tmp2 = Len(id & "!!??((") + tmp1
        If tmp1 > 0 Then
            getsubname = Mid(Str, tmp2, InStr(tmp1, Str, "##))==") - tmp2)
        Else
            getsubname = ""
		End If
	end function
	
	
	public Function detable(strHTML)
	  Dim objRegExp, strOutput
	  Set objRegExp = New Regexp	  
	  strOutput=strHTML	
	  objRegExp.IgnoreCase = True
	  objRegExp.Global = True
	  objRegExp.Pattern = "</?table[^>]*>"
	  strOutput = objRegExp.Replace(strOutput, "")	  
	  objRegExp.Pattern = "</?tr[^>]*>"
	  strOutput = objRegExp.Replace(strOutput, "")	  
	  objRegExp.Pattern = "</?td[^>]*>"
	  strOutput = objRegExp.Replace(strOutput, "")	  
	  objRegExp.Pattern = "</?th[^>]*>"
	  strOutput = objRegExp.Replace(strOutput, "")	
	  objRegExp.Pattern = "</?BLOCKQUOTE[^>]*>"
	  strOutput = objRegExp.Replace(strOutput, "")	  
	  objRegExp.Pattern = "</?tbody[^>]*>"
	  strOutput = objRegExp.Replace(strOutput, "")	  
	  objRegExp.Pattern = "<style[^\s]*"
	  strOutput = objRegExp.Replace(strOutput, "")	  
		detable = strOutput   
	  Set objRegExp = Nothing
	End Function
	
	public Function repl_domain(strHTML)
	  Dim objRegExp, strOutput
	  Set objRegExp = New Regexp	  
	  strOutput=strHTML	
	  objRegExp.IgnoreCase = True
	  objRegExp.Global = True
	  objRegExp.Pattern = "http://[^/.]*.vvblog.com/"
	  strOutput = objRegExp.Replace(strOutput, "http://www.vvblog.com/")	  
	  repl_domain = strOutput   
	  Set objRegExp = Nothing
	End Function
		
	public function profilthtm(strHTML)
		Dim objRegExp, strOutput
		Set objRegExp = New Regexp  
		strOutput=strHTML
		objRegExp.IgnoreCase = True
		objRegExp.Global = True	
		objRegExp.Pattern = "<img"
		strOutput = objRegExp.Replace(strOutput,"♂")		
		objRegExp.Pattern = "(♂[^>]*)>"
		strOutput = objRegExp.Replace(strOutput,"$1♀")		
		objRegExp.Pattern = "<[^>]*>"
		strOutput = objRegExp.Replace(strOutput,"")		
		objRegExp.Pattern = "style[^\s]*"
		strOutput = objRegExp.Replace(strOutput, "")		
		objRegExp.Pattern = "♂"
		strOutput = objRegExp.Replace(strOutput,"<img")		
		objRegExp.Pattern = "♀"
		strOutput = objRegExp.Replace(strOutput,">")		
		profilthtm = strOutput   
		Set objRegExp = Nothing
	end function
	
	public Function RemoveHTML(strHTML) 
		ON ERROR RESUME NEXT
		Dim objRegExp, strOutput
		  Set objRegExp = New Regexp
		  objRegExp.IgnoreCase = True
		  objRegExp.Global = True
		  objRegExp.Pattern = "<.+?>"
		  strOutput = objRegExp.Replace(strHTML, "")
		  strOutput = Replace(strOutput, "<", "<")
		  strOutput = Replace(strOutput, ">", ">")
		  RemoveHTML = strOutput   
		  Set objRegExp = Nothing
	End Function 

	public Function filtimg(strHTML)
	  Dim objRegExp, strOutput
	  Set objRegExp = New Regexp	  
	  strOutput=strHTML	
	  objRegExp.IgnoreCase = True
	  objRegExp.Global = True
	   if oblog.setup(28,0)=1 then
		objRegExp.Pattern="<img.[^>]*src(=| )(.[^>]*)>"
		strOutput = objRegExp.Replace(strOutput, "<img src=$2 onmousewheel=""return bbimg(this)"" >")
	  end if	   
	  if oblog.setup(27,0)>0 then
		'objRegExp.Pattern = "(<[^>]*src=\x22([^ ]*)\x22[^>]*>)"
		'strOutput = objRegExp.Replace(strOutput, "<a href=""$2"" target=""_blank"">$1</a>")
		
		'objRegExp.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
		'strOutput = objRegExp.Replace(strOutput,"<a target=""_blank"" href=$1>$1</a>")
		'objRegExp.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)$"
		'strOutput = objRegExp.Replace(strOutput,"<a target=""_blank"" href=$1>$1</a>")
		'objRegExp.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
		'strOutput = objRegExp.Replace(strOutput,"$1<a target=""_blank"" href=$2>$2</a>")
	
		'objRegExp.Pattern = "(<img[^>]*)>"
		 objRegExp.Pattern="<img.[^>]*src(=| )(.[^>]*)>"
	  	strOutput=objRegExp.replace(strOutput,"<img src=$2 onclick=""javascript:window.open(this.src);"" style=""CURSOR: pointer"">" )
		
		objRegExp.Pattern = "<img(.[^>]*)>"
		strOutput = objRegExp.Replace(strOutput,"<img$1 onload=""javascript:if(this.width>"&oblog.setup(27,0)&"){this.resized=true;this.style.width="&oblog.setup(27,0)&";}"">")
	  end if
	  	 
	  filtimg = strOutput   
	  Set objRegExp = Nothing
	End Function
	
	public Function filtskinpath(strHTML)
		ON ERROR RESUME NEXT         
		Dim objRegExp, strOutput
	 	 Set objRegExp = New Regexp  
	  	strOutput=strHTML
	 	 objRegExp.IgnoreCase = True
	 	 objRegExp.Global = True
		objRegExp.Pattern="src=([^\'^\u0022^\/^http^http\s^ftp^rt\sp^mm\s^#^\'^\.\.][A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
		strOutput=objRegExp.replace(strOutput,"src="""&blogurl&"$1""")
		objRegExp.Pattern="src=[\'\u0022]([^\/^http^http\s^ftp^rt\sp^mm\s^#^\'^\.\.][A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)[\'\u0022]"
		strOutput=objRegExp.replace(strOutput,"src="""&blogurl&"$1""")
		
		objRegExp.Pattern="href=([^\'^\u0022^\/^http^http\s^ftp^rt\sp^mm\s^#^\'^\.\.][A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
		strOutput=objRegExp.replace(strOutput,"href="""&blogurl&"$1""")
		objRegExp.Pattern="href=[\'\u0022]([^\/^http^http\s^ftp^rt\sp^mm\s^#^\'^\.\.][A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)[\'\u0022]"
		strOutput=objRegExp.replace(strOutput,"href="""&blogurl&"$1""")
		
		objRegExp.Pattern="url[\(]([^\/^http^http\s^ftp^rt\sp^mm\s^#^\'^\.\.][A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)[\)]"
		strOutput=objRegExp.replace(strOutput,"url("&blogurl&"$1)")
		
		objRegExp.Pattern="background=([^\'^\u0022^\/^http^http\s^ftp^rt\sp^mm\s^#^\'^\.\.][A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
		strOutput=objRegExp.replace(strOutput,"background="""&blogurl&"$1""")
		objRegExp.Pattern="background=[\'\u0022]([^\/^http^http\s^ftp^rt\sp^mm\s^#^\'^\.\.][A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)[\'\u0022]"
		strOutput=objRegExp.replace(strOutput,"background="""&blogurl&"$1""")
		
		objRegExp.Pattern="value=([^\'^\u0022^\/^http^http\s^ftp^rt\sp^mm\s^#^\'^\.\.][^# ]*\.[^# \'\u0022^]*)"
		strOutput=objRegExp.replace(strOutput,"value="""&blogurl&"$1""")
		objRegExp.Pattern="value=[\'\u0022]([^\/^http^http\s^ftp^rt\sp^mm\s^#^\'^\.\.][^# ]*\.[^# ]*)[\'\u0022]"
		strOutput=objRegExp.replace(strOutput,"value="""&blogurl&"$1""")
		if f_ext="asp" then
			strOutput=replace(strOutput,"<%","<％")
			'objRegExp.Pattern="\%\>"
			'strOutput=objRegExp.replace(strOutput,"％>")
			objRegExp.Pattern="(runat)[^>]*=[^>]*(server)"
			strOutput=objRegExp.replace(strOutput,"$1＝$2")
		end if
		filtskinpath=strOutput
		set objRegExp=nothing
	end Function
	public function filt_inc(strHTML)
	  ON ERROR RESUME NEXT
	  Dim objRegExp, strOutput
	  Set objRegExp = New Regexp  
	  strOutput=strHTML
	  objRegExp.IgnoreCase = True
	  objRegExp.Global = True
	  objRegExp.Pattern = "#include"
	  strOutput = objRegExp.Replace(strOutput, "＃i nclude")
	  objRegExp.Pattern = "#echo"
	  strOutput = objRegExp.Replace(strOutput, "＃e cho")
	  objRegExp.Pattern = "#flastmod"
	  strOutput = objRegExp.Replace(strOutput, "＃f lastmod")
	  objRegExp.Pattern = "#fsize"
	  strOutput = objRegExp.Replace(strOutput, "＃f size")
	  objRegExp.Pattern = "#exec"
	  strOutput = objRegExp.Replace(strOutput, "＃e xec")
	  objRegExp.Pattern = "#config"
	  strOutput = objRegExp.Replace(strOutput, "＃c onfig")
	  strOutput=replace(strOutput,"#此前在首页部分显示#","")
	   filt_inc = UBBCode(strOutput,1)   
	  Set objRegExp = Nothing
	end function
	public Function filt_include(strHTML)
	  Dim objRegExp, strOutput
	  Set objRegExp = New Regexp  
	  strOutput=strHTML
	  objRegExp.IgnoreCase = True
	  objRegExp.Global = True
	  objRegExp.Pattern = "<!-- #include file=[^>]*>"
	  strOutput = objRegExp.Replace(strOutput, oblog.setup(82,0))
	   filt_include = strOutput   
	  Set objRegExp = Nothing
	End Function
	
	public Function filtscript(V)
	  If Not Isnull(V) Then
			Dim t,test,Replacelist,t1,re,s,rnum
			Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			t=v
			t1=v
			s=v
			re.Pattern="&#36;"
			t1=re.Replace(t1,"$")
			re.Pattern="&#36"
			t1=re.Replace(t1,"$")
			re.Pattern="&#39;"
			t1=re.Replace(t1,"'")
			re.Pattern="&#39"
			t1=re.Replace(t1,"'")
			If InStr(str_htmlfilt,"|")=0 Then 
				Replacelist="(&#([0-9][0-9]*)|function|meta|window\.|object|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit=|error|click|key|load|focus|Blur))"
			Else
				Replacelist="("&str_htmlfilt&"&#([0-9][0-9]*)|function|meta|window\.|object|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit|error|click|key|load|focus|Blur))"
			End If
			re.Pattern="<((.[^>]*"&Replacelist&"[^>]*)|"&Replacelist&")>"
			Test=re.Test(t1)
			If Test=False Then
				If InStr(str_htmlfilt,"|")=0 Then 
					Replacelist="(&#([0-9][0-9]*)|function|meta|window\.|object|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit=|error|click|key|load|focus|Blur))"
				Else
					Replacelist="("&str_htmlfilt&"&#([0-9][0-9]*)|function|meta|window\.|object|script|js:|about:|file:|Document\.|vbs:|frame|cookie|on(finish|mouse|Exit|error|click|key|load|focus|Blur))"
				End If
				re.Pattern="(\[(.[^\]]*)\])((.[^\]]*"&Replacelist&"[^\]]*)|"&Replacelist&")(\[\/(.[^\]]*)\])"
				Test=re.Test(t1)
			End If
			If test Then
				Randomize
				rnum=cstr(Int(900*rnd)+1000)
				re.Pattern="\[(br)\]"
				s=re.Replace(s,"<$1>")
				re.Pattern = "(&nbsp;)"
				s = re.Replace(s,Chr(9))
				re.Pattern = "(<br>)"
				s = re.Replace(s,vbNewLine)
				re.Pattern = "(<p>)"
				s = re.Replace(s,"")
				re.Pattern = "(<\/p>)"
				s = re.Replace(s,vbNewLine)
				s=server.htmlencode(s)
				s="<div class=""quote""><strong>以下内容含脚本,或可能导致页面不正常的代码</strong><br /><TEXTAREA id=""scode"&rnum&""" style=""BORDER-RIGHT: 1px dotted; BORDER-TOP: 1px dotted; OVERFLOW-Y: visible; OVERFLOW: visible; BORDER-LEFT: 1px dotted; WIDTH: 98%; COLOR: #000000; BORDER-BOTTOM: 1px dotted"" rows=20 cols=120>"&s&"</TEXTAREA><br /><b>说明：</b>上面显示的是代码内容。您可以先检查过代码没问题，或修改之后再运行。<br /><input type=""button"" name=""run"" value=""运行代码"" onclick=""oblog_ViewCode('"&rnum&"');""></div>"
			end if
			filtscript=s
		End If
	End Function
	Private Function FixName(Byval UpFileExt)
		If IsEmpty(UpFileExt) Then Exit Function
		FixName = Lcase(UpFileExt)
		FixName = Replace(FixName,Chr(0),"")
		FixName = Replace(FixName,".","")
		FixName = Replace(FixName,"'","")
		FixName = Replace(FixName,"asp","")
		FixName = Replace(FixName,"asa","")
		FixName = Replace(FixName,"aspx","")
		FixName = Replace(FixName,"cer","")
		FixName = Replace(FixName,"cdx","")
		FixName = Replace(FixName,"htr","")
	End Function
	public Function UBBCode(strContent,CType)
		Dim re
		If CType=1 Then
			strContent = strContent
		Else
			strContent = strContent
		End If
		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=True
	
		If CType=1 Then
			'图片标签
			re.Pattern="\[IMG\](.[^\[]*)\[\/IMG\]"
			strContent=re.Replace(strContent,"<img src=""$1"" border=""0"">")
	
			'多媒体标签
			re.Pattern="\[MP=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/MP]"
			strContent=re.Replace(strContent,"<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3 width=$1 height=$2></embed></object>")
			re.Pattern="\[RM=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/RM]"
			strContent=re.Replace(strContent,"<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")
	
			re.Pattern="(\[FLASH\])(.[^\[]*)(\[\/FLASH\])"
			strContent= re.Replace(strContent,"<a href=""$2"" TARGET=_blank>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")
	
			re.Pattern="(\[FLASH=*([0-9]*),*([0-9]*)\])(.[^\[]*)(\[\/FLASH\])"
			strContent= re.Replace(strContent,"<a href=""$4"" TARGET=_blank>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE=""$4""><PARAM NAME=quality VALUE=high><embed src=""$4"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed></OBJECT>")
	
			're.Pattern="(\[sound\])(.[^\[]*)(\[\/sound\])" 
			'strContent=re.Replace(strContent,"<a href=""$2"" target=_blank><IMG SRC=images/files/mid.gIf border=0 alt='背景音乐'></a><bgsound src=""$2"" loop=""-1"">")
	
			re.Pattern="(\[URL\])(.[^\[]*)(\[\/URL\])"
			strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")
			re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
			strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$3</A>")
	
			re.Pattern="(\[EMAIL\])(\S+\@.[^\[]*)(\[\/EMAIL\])"
			strContent= re.Replace(strContent,"<A HREF=""mailto:$2"">$2</A>")
			re.Pattern="(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
			strContent= re.Replace(strContent,"<A HREF=""mailto:$2"" TARGET=_blank>$3</A>")
	
			're.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
			'strContent = re.Replace(strContent,"<a target=""_blank"" href=$1>$1</a>")
			're.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)$"
			'strContent = re.Replace(strContent,"<a target=""_blank"" href=$1>$1</a>")
			're.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@[\]\':+!]+)"
			'strContent = re.Replace(strContent,"$1<a target=""_blank"" href=$2>$2</a>")
	
			re.Pattern="\[color=(.[^\[]*)\](.[^\[]*)\[\/color\]"
			strContent=re.Replace(strContent,"<font color=$1>$2</font>")
			re.Pattern="\[face=(.[^\[]*)\](.[^\[]*)\[\/face\]"
			strContent=re.Replace(strContent,"<font face=$1>$2</font>")
			re.Pattern="\[align=(center|left|right)\](.*)\[\/align\]"
			strContent=re.Replace(strContent,"<div align=$1>$2</div>")
	
			re.Pattern="\[QUOTE\](.*)\[\/QUOTE\]"
			strContent=re.Replace(strContent,"<table style=""width:100%"" cellpadding=5 cellspacing=1><TR><TD style='border-right: #cccccc 1px solid; border-top: #cccccc 1px solid; border-left: #cccccc 1px solid; border-bottom: #cccccc 1px solid;background-color:#f6f6f6' width=""100%"">$1</td></tr></table><br>")
			re.Pattern="\[fly\](.*)\[\/fly\]"
			strContent=re.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
			re.Pattern="\[move\](.*)\[\/move\]"
			strContent=re.Replace(strContent,"<marquee scrollamount=3>$1</marquee>")		
			re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
			strContent=re.Replace(strContent,"<table width=$1 ><tr><td style=""filter:shadow(color=$2, strength=$3)"">$4</td></tr></table>")
			re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
			strContent=re.Replace(strContent,"<table width=$1 ><tr><td style=""filter:glow(color=$2, strength=$3)"">$4</td></tr></table>")
	
			re.Pattern="\[i\](.[^\[]*)\[\/i\]"
			strContent=re.Replace(strContent,"<i>$1</i>")
			re.Pattern="\[u\](.[^\[]*)(\[\/u\])"
			strContent=re.Replace(strContent,"<u>$1</u>")
			re.Pattern="\[b\](.[^\[]*)(\[\/b\])"
			strContent=re.Replace(strContent,"<strong>$1</strong>")
			re.Pattern="\[size=([1-4])\](.[^\[]*)\[\/size\]"
			strContent=re.Replace(strContent,"<font size=$1>$2</font>")
			strContent=replace(strContent,"<I></I>","")
		End If
	
		If CType=1 Or CType=2 Then
			're.Pattern="\[UPLOAD=(gif|jpg|jpeg|bmp|png)\](.[^\[]*)(gif|jpg|jpeg|bmp|png)\[\/UPLOAD\]"
			'strContent= re.Replace(strContent,"<a href=""$2$1"" TARGET=_blank><IMG SRC=""$2$1"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></a>")
	
			re.Pattern="\[UPLOAD=(.[^\[]*)\](.[^\[]*)\[\/UPLOAD\]"
			strContent= re.Replace(strContent,"<br />$2<a href=""$1"">(点击浏览该文件)</a>")
		End If
		'If InStr(strContent,"[EA_AdRotator]")<>0 Then
		'	strContent=Replace(strContent,"[EA_AdRotator]","<table align=left><tr><td>"&AdRotator(Site_Info(17),Chr(10))&"</td></tr></table>")
		'End If
		Set re=Nothing
		UBBCode=strContent
	End Function
	
	'过滤掉flash UBB标记
	'[flash=500,350]http://www.kunfu.net/movie.swf[/flash]
	Function FilterUBBFlash(byval strFlash)
		Dim strFlash1
		strFlash1=LCase(strFlash)
		If InStr(strFlash1,"[/flash]")>0 Then
			strFlash1 = Replace(strFlash1,"[/flash]","[ /flash ]")
			strFlash1 = Replace(strFlash1,"[flash","[ flash ")
			FilterUBBFlash=strFlash1
		Else
			FilterUBBFlash=strFlash
		End If
	End Function
	
	Public Function trimlog(logtext, showword)
        On Error Resume Next
        Dim Contentlen
        If InStr(logtext, "#此前在首页部分显示#") > 0 Then
            trimlog = Left(logtext, InStrRev(logtext, "#此前在首页部分显示#") - 1) & "<br />……"
            Exit Function
        End If
        If IsNull(showword) Or showword = "" Then showword = 0
        Contentlen = oblog.strLength(logtext)
        If Contentlen <= showword Or showword = 0 Then
            trimlog = logtext
            Exit Function
          Else
            If InStrRev(logtext, "<object") > 0 Or InStrRev(logtext, "<OBJECT") > 0 Then
                If showword < 100 Then
                    trimlog = ""
                  Else
                    trimlog = detable(logtext)
                End If
              Else
                trimlog = oblog.InterceptStr(detable(logtext), showword + 100)
                If InStrRev(trimlog, "<P", -1, 1) > 0 And (Len(trimlog) - InStrRev(trimlog, "<P", -1, 1)) < 400 Then
                    trimlog = Left(trimlog, InStrRev(trimlog, "<P", -1, 1) - 1)
                  ElseIf InStrRev(trimlog, "<img", 1) > 0 And (Len(trimlog) - InStrRev(trimlog, "<img", 1)) < 400 Then
                    trimlog = Left(trimlog, InStrRev(trimlog, "<img", 1) - 1)
                  ElseIf InStrRev(trimlog, "。") > 0 And (Len(trimlog) - InStrRev(trimlog, "。")) < 400 Then
                    trimlog = Left(trimlog, InStrRev(trimlog, "。"))
                  ElseIf InStrRev(trimlog, "<br", 1) > 0 And (Len(trimlog) - InStrRev(trimlog, "<br", 1)) < 400 Then
                    trimlog = Left(trimlog, InStrRev(trimlog, "<br", 0, 1) - 1)
                    'elseif Instrrev(trimlog,"<object",-1,1) > 0 and (Len(trimlog) - Instrrev(trimlog,"<object",-1,1))< 200 then
                    '   trimlog = Left(trimlog,InstrRev(trimlog,"<object",-1,1)-1)
                  ElseIf InStrRev(trimlog, "?") > 0 And (Len(trimlog) - InStrRev(trimlog, "?")) < 400 Then
                    trimlog = Left(trimlog, InStrRev(trimlog, "?"))
                End If
            End If
            'if instr(1,trimlog,"<object",1)<>0 then    trimlog=left(detable(logtext),instr(1,detable(logtext),"</object>",1)+9-1)
            trimlog = trimlog & "<br/>……"
        End If
    End Function
	
%>