
<script>
function menu(oblog){
if(oblog.style.display=='')oblog.style.display='none';
else oblog.style.display='';
}
</script>

<style>
	img {
	background:#ffffff;
	border:0px;
	border-collapse: collapse;
	font-size:13px;
	}
	.menu-t1 {
	/*background:#eeeeee;*/
	font-size:13px;
	cursor:hand;
margin:6px,3px,6px,0px;
	padding-top:6px;
	}
</style>
<div style="position:absolute; top:250px; left:830px; width:120px;">
	<img src="/images/mailmanage.gif"><a href="HomestayBackctrl-personalMessageCooperate.asp">��ҵ��Ϣ(<%=chkpm()%>)</a>
</div>

<p class="public-title011">
	<%=oblog.logined_uname%>�������û�����
</p>
	<div id="toplogo"></div>
	<div id="sub">
	<img src='/images/[MyHomestay]Manager.gif' /><a href="HomestayBackctrl-indexCooperate.asp" target="_top">������ҳ</a>| <img src='/images/logout.gif' /><a href="HomestayBackctrl-indexCooperate.asp?t=logout&re=1">�˳���̨</a>
	 <!--<br />
	 <img src="/images/mailmanage.gif"><a href="HomestayBackctrl-personalMessageCooperate.asp">��ҵ��Ϣ(<%=chkpm()%>)</a>--><br />&nbsp;
	</div>
	<div class="menu-t1" onClick="menu(oblog_1)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> �ύ�������</div>
	<ul id="oblog_1">
	
	<li><a href="HomestayBackctrl-PostSubmitManageCooperate.asp" title="�鿴�����ύ�����������"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> ������Ϲ���</a></li>
	<li><a href="HomestayBackctrl-PostSubmitManageCooperate.asp?usersearch=5" title="�ݴ��ڲݸ����У����ύ��"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> �ݴ�ݸ���</a></li>
	
	<% If oblog.logined_uisValidate=0 Then %>
	<li><a  title="�����ȵȴ�ͨ�����..." onclick="alert('�����ȵȴ�ͨ�����...');" style="cursor:hand;"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> <font color="#038ad7">�ύ�����Ϣ</font></a></li>
	<% Else %>
	<li><a href="HomestayBackctrl-PostSubmitCooperate.asp" title="���ǽ�����24Сʱ�ڣ�������ύ���ϵ���˹��������価�췢����ס�����ϡ�"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> <font color="#038ad7">�ύ�����Ϣ</font></a></li>
	<% End If %>
	
	</ul>
	
	
	
	
	
	<div class="menu-t1" onClick="menu(oblog_2)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> �����û����Ϲ���</div>
	<ul id="oblog_2">
	
	
	<li><a href="HomestayBackctrl-InfoCooperate.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/11.gif" align="absmiddle" /> �޸��ҵļ�������</a></li>
	
	<li><a href="HomestayBackctrl-PassSettingCooperate.asp?action=userpassword&pass=1"><img src="oBlogStyle/OblogStyleAdminImages/li/14.gif" align="absmiddle" /> �޸��ҵ�����</a></li>
	<li><a href="HomestayBackctrl-PassSettingCooperate.asp?action=userpassword"><img src="oBlogStyle/OblogStyleAdminImages/li/14.gif" align="absmiddle" /> ����������ʾ</a></li>
	
	</ul>

	
	
	