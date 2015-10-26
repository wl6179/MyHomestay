<!--#include file="inc/inc_syssite.asp"-->
<!--#include file="inc/syscode.asp"-->
<%

'*********************************************************
'File: 			photoPlayer.asp
'Description:	相册列表自动播放器 For oBlog3.1
'Author: 		王旭 
'Copyright:		http://www.myhomestay.com.cn
'LastUpdate:	20050921
'*********************************************************

If EN_PHOTO=0 then
	oblog.adderrstr("此功能已被系统关闭！")
	oblog.showerr
	Set oblog = Nothing
End If

'--------------------------------------------
'支持以下几种参数
'网站最新图片
'PhotoPlayer.asp
'系统分类
'PhotoPlayer.asp?classid=1
'用户最新图片
'PhotoPlayer.asp?userid=1
'用户分类
'PhotoPlayer.asp?userid=1&subject=1
'--------------------------------------------

Dim classid,userid,subjectid,sql,rst
Dim picNumber,ImgJS,i
classid=Request("classid")
userid=Request("userid")
subjectid=Request("subjectid")

sql="Select Top 100 * From oblog_upfile Where isPower=0 And isPhoto=1 "
'Response.Write Now()
If classid<>"" Then
	If IsNumeric(classid) Then
		classid=CLNG(classid)
		sql= sql & " And sysClassid=" & classid
	Else
		classid=""
	End If
End If
If userid<>"" Then
	If IsNumeric(userid) Then
		userid=CLNG(userid)
		sql= sql & " And userid=" & userid
	Else
		userid=""
	End If
End If
If subjectid<>"" Then
	If IsNumeric(subjectid) Then
		subjectid=CLNG(subjectid)
		sql= sql & " And userClassid=" & subjectid
	Else
		subjectid=""
	End If
End If
sql = sql & " order By fileid Desc"
Set rst= Server.CreateObject("Adodb.RecordSet")

Call link_database
rst.Open sql,conn,1,3
If rst.Eof Then
	Set rst=Nothing
	Response.Write "没有任何符合条件的图片"
	conn.close
	Set conn=Nothing
	Response.End
Else
	picNumber=rst.Recordcount
	ImgJS="ImgName[0]=""" & blogdir & "images/PhotoConver.gif"";" & VBCRLF
	For i=1 To picNumber
  		ImgJS = ImgJS & "ImgName[" & i & "]=""" &  rst("file_path") & """;" & VBCRLF
  		rst.MoveNext
  	Next
End If
rst.Close
Set rst=Nothing
conn.close
Set conn=Nothing
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>MyHomestay图片自动播放器</title>
<script language="JavaScript">
var goid=0;
var t=0;

  function ImgArray(len)
  {
   this.length=len;
   }
  ImgName=new ImgArray(<%=picNumber%>);
  <%
  Response.write imgJS
  %>
  function playImg()
  {
    t_end=document.player.intsec.value;
	if (t==<%=picNumber%>)
	   {
	    t=0;
	   }
	   else
	   {t++;}
	if (t==<%=picNumber%>)
	   {
	    t_end=10000;
	   }
	if (goid==0){
	  img.style.filter="blendTrans(Duration=1)";
	  img.filters[0].apply();
	  img.src=ImgName[t];
	  tIndex=t;
	  img.filters[0].play();
	  mytimeout=setTimeout("playImg()",t_end);
	}
   }
   
   function go(id){
   	if (id==1){
   		goid=0
   		playImg();
   		}
   	else if(id==2){
   		goid=1;
   		}
   	else if(id==3){ 
   		goid=1;
   		t=t+1;
   		if (t<=<%=picNumber%>) img.src=ImgName[t+1]; 		
	}
	else if(id==4){
		goid=1;
		t=t-1;
		 if (t>0) img.src=ImgName[t];
	}
	else if(id==5){
		goid=1;
		t=0;
		playImg();
	}
	else{
		goid=0;
		t=0;
		playImg();
		}
		
   	}
</script>

<style>
body img{border:3px ridge gold; }
td,body {
	color:#4b4b4b;
	font-size:12px;
	font-family:arial,sans-serif;	
	}
.photobox {	
	border-top: 2px dotted #C6CFBD;
	border-left: 2px dotted #C6CFBD;
	border-right: 2px dotted #C6CFBD;
	border-bottom: 2px dotted #C6CFBD;
}

</style>


</head>

<body bgcolor="#FFFFFF"  onload="playImg()">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="458">
  	 <tr>
      <td width="458" align=center>
      <p align="center"><b><font color="#000000">oBlog3.1 图片自动播放器</font></b><BR/><BR/></td>
    </tr>
  	<form name="player" method="post" action="">
  	<tr>
  	<td>
  		播放频率：
  			<select name="intsec">
  				<option value=1000>1秒</option>
  				<option value=3000 Selected>3秒</option>
  				<option value=5000>5秒</option>
  				<option value=8000>8秒</option>
  				<option value=10000>10秒</option>
  			</select>
			<input type="button" value="开始" onClick="javascipt:go(1)">
  			<input type="button" value="停止" onClick="javascipt:go(2)">
  			<input type="button" value="上一张" onClick="javascipt:go(3)">
  			<input type="button" value="下一张" onClick="javascipt:go(4)">
  			<input type="button" value="关闭" onClick="javascipt:window.close();">
  			<BR/>
  			<BR/>
  	</td>
  	</tr>
	</form>
   <tr>
      <td width="458" class="photobox" align=center valign=middle>
      	<BR/>
      <center><img src="<%=blogdir%>images/photoconver.gif" name="img" onload="if(this.width>=400)this.width=400;" onClick="window.open(img.src);" style="cursor:hand"></center>
      	<BR/>
    </tr>
  </table>
  </center>
</div>
<p>&nbsp;</p>

</body>
</html>