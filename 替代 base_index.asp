<!--#include file="inc/inc_sys.asp"-->
<%
select case request("action")
case "top"
	call admin_top()
case "left"
	call admin_left()
case "main"
	call admin_main()
case "state"
	If Application(cache_name_user&"_systemstate")<>"stop" Then
		Application(cache_name_user&"_systemstate")="stop"
	Else
		Application(cache_name_user&"_systemstate")="run"
	End If
	Application(cache_name_user&"_systemnote")=Request.Form("systemnote")
	Response.Write "<script language=javascript>parent.location.href=""admin_index.asp"";</script>"
case else
	call main()
end select

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

sub main()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ҵ�ס���� MyHomestay--��̨����</title>
</head>
<frameset rows="*" cols="251,*" framespacing="0" frameborder="0" border="false" id="frame" scrolling="yes">
  <frame name="left" scrolling="auto" marginwidth="0" marginheight="0" src="admin_index.asp?action=left">
  <frameset framespacing="0" border="false" rows="35,*" frameborder="0" scrolling="yes">
    <frame name="top" scrolling="no" src="admin_index.asp?action=top">
    <frame name="main" scrolling="auto" src="admin_index.asp?action=main">
  </frameset>
</frameset>
<noframes>
  <body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <p>���������汾���ͣ���ϵͳҪ��IE5�����ϰ汾����ʹ�ñ�ϵͳ��</p>
  </body>
</noframes>
</html>
<%
end sub

sub admin_top()
%>
<html>
<head>
<title>�ҵ�ס���� MyHomestay ��̨����ҳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
a:link {
	color:#333333;
	text-decoration:none;
	font-size: 12px;
}
a:hover {color:#2786CB;}
a:visited {color:#333333;text-decoration:none}

td {FONT-SIZE: 9pt;COLOR: #000000; FONT-FAMILY: "����"}
img {filter:Alpha(opacity:100); chroma(color=#FFFFFF)}
body {background:#ffffff;}
</style>
<base target="main">
<script>
function preloadImg(src)
{
	var img=new Image();
	img.src=src
}
preloadImg("images/admin/admin_top_open.gif");

var displayBar=true;
function switchBar(obj)
{
	if (displayBar)
	{
		parent.frame.cols="0,*";
		displayBar=false;
		obj.src="images/admin/admin_top_open.gif";
		obj.title="����߹���˵�";
	}
	else{
		parent.frame.cols="250,*";
		displayBar=true;
		obj.src="images/admin/admin_top_close.gif";
		obj.title="�ر���߹���˵�";
	}
}
</script>
</head>
<body leftmargin="0" topmargin="0" style="background:#ffffff;">
<table width="100%" height="100%" border=0 cellpadding=0 cellspacing=0 bgcolor="#ffffff">
  <tr valign=middle> 
    <td width=10> <img onClick="switchBar(this)" src="images/admin/admin_top_close.gif" title="�ر���߹���˵�" style="cursor:hand"> </td>
	<td width=80><a href="admin_adminmodifypwd.asp"><font style="font-size:14px;"><b>�޸�����</b></font></a><img src='/images/loader.gif' /></td>
    <td align="left" width="500">�ҵ�ס���� MyHomestay רҵ��ס����̷�����&nbsp;&nbsp;&nbsp;&nbsp;�ͻ����ߣ�010-85493388/13910652850</td>
    <td width="50" align="left"><a href="../index.asp" target="_blank">վ����ҳ</a></td>
  </tr>
</table>
</body>
</html>
<%end sub
sub admin_left()
%>
<html>
<head>
<title>MyHomestay--������</title>
<STYLE type=text/css>
BODY {
 BACKGROUND: #ffffff url(images/admin/topbg3.gif) repeat-x fixed top; MARGIN: 0px; FONT: 12px ����;
 	scrollbar-face-color:#ffffff;
	scrollbar-highlight-color: #cccccc;
	scrollbar-shadow-color: #cccccc;
	scrollbar-3d-light-color: #ffffff;
	scrollbar-arrow-color : #cccccc;

	SCROLLBAR-TRACK-COLOR: #ffffff;
	SCROLLBAR-DARKSHADOW-COLOR: #ffffff;
}
TABLE {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
TD {
	FONT: 12px ����
}
IMG {
	BORDER-RIGHT: 0px; BORDER-TOP: 0px; VERTICAL-ALIGN: bottom; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px
}
A {
	FONT: 12px ����; COLOR: #333333; TEXT-DECORATION: none
}
A:hover {
	COLOR: #2786CB; TEXT-DECORATION: underline
}
.sec_menu {
	BORDER-RIGHT: white 1px solid; BACKGROUND: #BFE0FB; OVERFLOW: hidden; BORDER-LEFT: white 1px solid; BORDER-BOTTOM: white 1px solid
}
.menu_title {
	
}
.menu_title SPAN {
	FONT-WEIGHT: normal;
	LEFT: 8px;
	COLOR: #333333;
	POSITION: relative;
	TOP: 2px;
	font-size: 14px;
}
.menu_title2 {
	
}
.menu_title2 SPAN {
	FONT-WEIGHT: normal;
	LEFT: 8px;
	COLOR: #333333;
	POSITION: relative;
	TOP: 2px;
	font-size: 14px;
}
</STYLE>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width=100% cellpadding=0 cellspacing=0 border=0 align=left>
    <tr>
  <td valign=top class="menu_title2"> 
    <table width=158 border="0" align=center cellpadding=0 cellspacing=0>
  <tr>
    <td height=42 valign=bottom>
	  <img src="../images/homestay-s.gif">
    </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/admin/bg.gif id=menuTitle0> 
            <span><a href="admin_index.asp?action=main" target=main><b>������ҳ</b></a> 
            | <a href="admin_login.asp?action=logout" target="_top" ><b>�˳�</b></a></span> </td>
  </tr>
  <tr>
    <td style="display:" id='submenu0'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
	<tr>
                
              </tr>
    <tr>
                
              </tr>
<tr>
                <td height=20>�û�����<%= session("adminname") %></td>
              </tr>
<tr>
   
              </tr>
</table>
</div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
</div>
	</td>
  </tr>
</table>


<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle onClick="showsubmenu(23)" style="cursor:hand;"> 
            <span>���·��� + ��������</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu23'>
<div class=sec_menu style="width:158">
            <table width=130 align=center cellpadding=0 cellspacing=0 >
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_logclass.asp" target="main">���·������</a></td>
              </tr>
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_UserLogclass_Nav.asp target=main>��Ŀ�ռ����+����</a></td>
              </tr>
			  <!--<tr>
                <td height=20 ><img src="images/admin/bullet.gif"></td>
                <td height=20 ><a href="admin_user.asp?action=gouser1" target="main">ֱ�ӽ�����Ŀ����</a></td>
              </tr>-->
              
              <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_UserLogclass_Nav.asp?Action=gouser2&username=myhomestay"  target=main >����վ����Ϣ</a></td>
              </tr>
			   <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=admin_pmall.asp  target=main >����վ����Ϣ</a></td>
              </tr>
            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>


<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle4 onClick="showsubmenu(4)" style="cursor:hand;"> 
            <span>�û�����</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu4'>
<div class=sec_menu style="width:158">
            <table width=130 align=center cellpadding=0 cellspacing=0 >
              
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_user.asp target=main>�й���ͥ�û�����</a></td>
              </tr>
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_userCooperate.asp target=main>�����û�����</a></td>
              </tr>
      
      		<tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=admin_admin.asp?Action=Add target="main">��ӹ���Ա</a> | <a href=admin_admin.asp target=main>����</a></td>
              </tr>
			  
				<!--<tr>
                <td height=20 ><img src="images/admin/bullet.gif"></td>
                <td height=20 ><a href="admin_blogstar.asp" target="main">�û�����</a></td>
              </tr>-->
              

              <!--<tr>
                <td height=20 ><img src="images/admin/bullet.gif"></td>
                <td height=20 ><a href="admin_rename.asp" target="main">�û�����</a></td>
              </tr>-->
              
              

            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>




<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title9'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle9 onClick="showsubmenu(9)" style="cursor:hand;"> 
            <span>������Ϲ���</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu9'>
<div class=sec_menu style="width:158">
            <table width=130 align=center cellpadding=0 cellspacing=0 >
       
    
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_userCooperate_PassTeacherNo.asp target=main>�鿴������������</a></td>
              </tr>
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_userCooperate_PassTeacher.asp target=main>�鿴������������</a></td>
              </tr>
              
              <tr>
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_userCooperate_PassTeacherDeleted.asp target=main>��ɾ�����������</a></td>
              </tr>
              

            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>




<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle2 onClick="showsubmenu(2)" style="cursor:hand;"> 
            <span>��������</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu2'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>



	      
              <!--<tr> 
                <td width="15" height=20 align="right"><img src="images/admin/bullet.gif"></td>
                <td width="115"><a href=admin_setup.asp target=main>��վ��Ϣ����</a></td>
              </tr>-->
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_userclass.asp" target="main">�û��������</a></td>
              </tr>
              
              <!--<tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_logclass.asp?t=1" target="main">���������</a></td>
              </tr>-->
              <!--<tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_tags.asp" target="main">ϵͳTAG����</a></td>
              </tr>-->
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=admin_filtrate.asp target=main>���������ֹ��˹���</a></td>
              </tr>
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=admin_lockip.asp target=main>����IP����</a></td>
              </tr>
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=1 target=main>�޸���������</a> | <a href=admin_note.asp?action=do1 target=main>�ı�</a></td>
              </tr>
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=3 target=main>�޸�ע������</a> | <a href=admin_note.asp?action=do3 target=main>�ı�</a></td>
              </tr>
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=2 target=main>�޸���վ����</a> | <a href=admin_note.asp?action=do2 target=main>�ı�</a></td>
              </tr>
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=admin_ad.asp target=main>�û�ҳ�������</a></td>
              </tr>
              <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href=../admin_edit.asp?do=4 target=main>�û���̨֪ͨ</a> | <a href=admin_note.asp?action=do4 target=main>�ı�</a></td>
			  </tr>
              <!--<tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_userdir.asp" target="main">�û�Ŀ¼����</a></td>
              </tr>-->
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_count.asp" target="main">����ϵͳ����</a></td>
              </tr>
              
              <tr>
                <td height=20 ><img src="images/admin/bullet.gif"></td>
                <td height=20 ><a href="admin_user.asp?Action=Update" target="main">�����û���̬ҳ</a></td>
              </tr>
              
            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>

<!--<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle5 onClick="showsubmenu(5)" style="cursor:hand;"><span>ģ�����</span>
          </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu5'><div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_sysskin.asp?action=addskin target=main>���ϵͳģ��</a> 
                  | <a href=admin_sysskin.asp?action=showskin target=main>����</a></td>
              </tr>
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_skin.asp?action=insys1" target="main">ϵͳģ�浼��</a> 
                  | <a href="admin_skin.asp?action=outsys" target="main">����</a></td>
              </tr>
              <tr> 
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_userskin.asp?action=addskin" target="main">����û�ģ��</a> 
                  | <a href="admin_userskin.asp?action=showskin&ispass=1" target="main">����</a></td>
              </tr>
              <tr>
                <td height=20><img src="images/admin/bullet.gif"></td>
                <td height=20><a href="admin_skin.asp?action=inuser1" target="main">�û�ģ�浼��</a> 
                  | <a href="admin_skin.asp?action=outuser" target="main">����</a></td>
              </tr>
            </table>
      </div>
        <div  style="width:158">
          <table cellpadding=0 cellspacing=0 align=center width=130>
            <tr>
              <td height=20 class="menu_title"></td>
            </tr>
          </table>
      </div></td>
  </tr>
</table>-->
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle5 onClick="showsubmenu(6)" style="cursor:hand;"> 
            <span>�ϴ��ļ�����</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu6'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_uploadfile_user.asp target=main>�ϴ������û��嵥</a><a href=admin_sysskin.asp?action=showconfig target=main></a></td>
              </tr>
				<tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_uploadfile.asp target=main>�ϴ������ļ��嵥</a><a href=admin_sysskin.asp?action=showconfig target=main></a></td>
              </tr>

            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20 class="menu_title"></td>
              </tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%
'�����SQL Server���ݿ�������ʾ�ý�
If IS_SQLDATA=0 Then
%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle7 onClick="showsubmenu(7)" style="cursor:hand;"> 
            <span>���ݿ����</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu7'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td width="15" height=20><img src="images/admin/bullet.gif"></td>
  <td><a href=admin_database.asp?Action=Compact target=main>ѹ�����ݿ�</a></td>
</tr>
<tr><td height=20><img src="images/admin/bullet.gif"></td>
  <td height=20><a href=admin_database.asp?Action=Backup target=main>�������ݿ�</a></td>
</tr>
<tr><td height=20><img src="images/admin/bullet.gif"></td>
  <td height=20><a href=admin_database.asp?Action=Restore target=main>�ָ����ݿ�</a></td>
</tr>
<tr><td height=20><img src="images/admin/bullet.gif"></td>
  <td height=20><a href=admin_database.asp?Action=SpaceSize target=main>ϵͳ�ռ�ռ��</a></td>
</tr>
</table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%End If%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle5 onClick="showsubmenu(8)" style="cursor:hand;"> 
            <span>ϵͳ����</span> </td>
  </tr>
  <tr>
    <td style="display:display" id='submenu8'>
<div class=sec_menu style="width:158">
            <table cellpadding=0 cellspacing=0 align=center width=130>
              <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_analyze.asp?action=year target=main>������ȷ���</a></td>
              </tr>
				<tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_analyze.asp?action=month target=main>�����¶ȷ���</a></td>
              </tr>
			  <tr> 
                <td width="15" height=20><img src="images/admin/bullet.gif"></td>
                <td><a href=admin_analyze.asp?action=hour target=main>����ʱ�η���</a></td>
              </tr>

            </table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr>
                <td height=20 class="menu_title"></td>
              </tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
          <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/bg.gif" id=menuTitle9> 
            <span>MyHomestay��Ϣ</span> </td>
  </tr>
  <tr>
    <td>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=130>
<tr><td height=20><br>
                  ��Ȩ<a href="http://www.myhomestay.com.cn" target=_blank>myhomestay.com.cn</a><br>
                  ���������&nbsp;<a href='mailto:wangliang6179@163.com'>����</a><BR>
                  ����֧�֣�&nbsp;<a href="http://www.myhomestay.com.cn" target=_blank>�ҵ�ס����</a><br>                  
                 
                  <br>
</td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
</table>
</body>
</html>
<%
'response.Write("<script src=""http://oblog.myhomestay.com.cn/count.asp?a="&oblog.setup(3,0)&"&b="&oblog.setup(4,0)&"&c="&oblog.setup(21,0)&"&d="&oblog.ver&"&e="&is_sqldata&"&f="&oblog.setup(6,0)&"&g="&oblog.setup(24,0)&"""></script>")
end sub
sub admin_main()
Dim theInstalledObjects(17)
theInstalledObjects(0) = "MSWC.AdRotator"
theInstalledObjects(1) = "MSWC.BrowserType"
theInstalledObjects(2) = "MSWC.NextLink"
theInstalledObjects(3) = "MSWC.Tools"
theInstalledObjects(4) = "MSWC.Status"
theInstalledObjects(5) = "MSWC.Counters"
theInstalledObjects(6) = "IISSample.ContentRotator"
theInstalledObjects(7) = "IISSample.PageCounter"
theInstalledObjects(8) = "MSWC.PermissionChecker"
theInstalledObjects(9) = "Scripting.FileSystemObject"
theInstalledObjects(10) = "adodb.connection"
    
theInstalledObjects(11) = "SoftArtisans.FileUp"
theInstalledObjects(12) = "SoftArtisans.FileManager"
theInstalledObjects(13) = "JMail.SMTPMail"
theInstalledObjects(14) = "CDONTS.NewMail"
theInstalledObjects(15) = "Persits.MailSender"
theInstalledObjects(16) = "LyfUpload.UploadFile"
theInstalledObjects(17) = "Persits.Upload.1"
%>
<html>
<head>
<title>MyHomestay��̨������ҳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="images/admin/admin_STYLE.CSS">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">
<br>
<table cellpadding="2" cellspacing="1" border="0" width="98%" class="border" align=center>
<tr align="center">
    <td width="300" height=25 class="topbg"><div align="left"><strong>MyHomestay��̨������ҳ</strong></div>
</tr>
</table>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center"> 
<td height=25 colspan=2 class="topbg"><strong>ϵͳ������ر�</strong></th></tr> 
     <tr> 
    <td width="20%" class="tdbg" height=23>��ǰϵͳ״̬</td>
    <td width="80%" class="tdbg">&nbsp;&nbsp;<font color=green>
    <%
    	Dim strSubmit
    	If Application(cache_name_user&"_systemstate")="stop" Then
    		Response.Write "<font color=red>�ر���</font>"
    		strSubmit="��������"
    	Else
    		Response.Write "����������..."
    		strSubmit="�ر�ϵͳ"
    	End If
    %></font>
    </td>
  </tr>
   <form name="systemcontrol" method="post" action="admin_index.asp?action=state">
	  <tr> 
	    <td width="20%" class="tdbg" height=23>�ر�ʱ��ʾ��</td>
	    <td width="80%" class="tdbg"> 
	    	<table width="100%" border="0" cellspacing="0" cellpadding="2">
	        <tr> 
	          <td>
	          	<textarea cols=60 rows=5 name="systemnote"><%
	          		If Application(cache_name_user&"_systemnote")="" Then
	          			Response.Write "ϵͳ����ά�����Ժ󿪷�"
	          		Else
	          			Response.Write Application(cache_name_user&"_systemnote")
	          		End If
	          		%>
	          		</textarea>	          	
	          	</td>
	        </tr>
	      </table></td>
	  </tr>
	  <tr>
	    <td class="tdbg" height=23></td>
	    <td class="tdbg"><input type="submit" value="<%=strSubmit%>">(ϵͳ���ڹر�״̬ʱ,�Թ���Ա���������������)</td>
	  </tr>
	</form>

</table>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center">
    <td height=25 class="topbg"><strong>MyHomestay �� ��</strong>
  <tr>
    <td height=23 class="tdbg">1��<a href="../jsreadme.asp" target="_blank"><strong>��ҳ�����ļ��Ĳ��Լ�����˵����������</strong>��</a></td>
  </tr>
  <tr>
    <td height=23 class="tdbg">2��<a href="admin_skin_help.asp" target="_blank"><strong>ϵͳģ�漰�û�ģ��ı��˵����������</strong>��</a></td>
  </tr>
  <tr>
    <td height=23 class="tdbg"><p>3���û�Ȩ��<br> 
        ���ٴ�����û������ܷ������¡�<br>      
        ����ע���û������Է������¡�<br>
      ����VIP�û�������ͨע���û�������Ȩ�ޣ��ɵ��������ϴ��ռ䡣<br>
      ����ǰ̨����Ա����VIP�û�������Ȩ�ޣ�������޸ģ�ɾ�������û������¡�<br>
      ���ݸ��û�����ϴ�Ȩ�ޣ�����վ��Ϣ�������趨��</p>
    </td>
  </tr>
  <tr>
    <td class="tdbg">4�������������ʻ������Ժ󣬴��ʻ�����������ҳ�潫�����Ρ�</td>
  </tr>

  <tr>
    <td height=23 class="tdbg">5�����ϴ��ļ��������������Ƿ��ļ��ߴ���󼰷������Ƿ�֧��fso��</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">6�����κ����⣬����ѯMyHomestay�ٷ���վ<a href="http://www.myhomestay.com.cn" target="_blank">http://www.myhomestay.com.cn</a>��</td>
  </tr>
</table>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center"> 
    <td height=25 colspan=2 class="topbg"><strong>�� �� �� �� Ϣ</strong><tr> 
  <tr> 
    <td width="50%"  class="tdbg" height=23>���������ͣ�<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
    <td width="50%" class="tdbg">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>վ������·����<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
    <td width="50%" class="tdbg">���ݿ��ַ��</td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>FSO�ı���д��
      <%If Not IsObjInstalled(theInstalledObjects(9)) Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>

      <%end if%></td>
    <td width="50%" class="tdbg">���ݿ�ʹ�ã�
      <%If Not IsObjInstalled(theInstalledObjects(10)) Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>
      <%end if%></td>
  </tr>
  <tr> 
    <td width="50%" class="tdbg" height=23>Jmail���֧�֣�
      <%If Not IsObjInstalled(theInstalledObjects(13)) Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>
      <%end if%></td>
    <td width="50%" class="tdbg">CDONTS���֧�֣�
      <%If Not IsObjInstalled(theInstalledObjects(14)) Then%>
      <font color="red"><b>��</b></font>
      <%else%>
      <b>��</b>
      <%end if%></td>
  </tr>
  <tr>
    <td class="tdbg" height=23>&nbsp;</td>
    <td align="right" class="tdbg">&nbsp;</td>
  </tr>
</table>
<br>
<table width="98%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center"> 
    <td height=25 colspan=2 class="topbg"><strong>MyHomestay����</th></strong> 
  <tr> 
    <td width="20%" class="tdbg" height=23>���ĳ�����ƣ�</td>
    <td width="80%" class="tdbg"> <table width="100%" border="0" cellspacing="0" cellpadding="2">
        <tr> 
          <td width="60">����</td>
          <td>��&nbsp;Email��wangliang6179@163.com;&nbsp;&nbsp;&nbsp;&nbsp;QQ��4985612��</td>
        </tr>
      </table></td>
  </tr>


  <tr>
    <td class="tdbg" height=23>������ƣ�</td>
    <td class="tdbg">����</td>
  </tr>
</table>
</body>
</html>
<%

end sub
%>

