
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
	<a href="user_pmmanage.asp"><img src="/images/mailmanage.gif">ס������Ϣ(<%=chkpm()%>)</a>
</div>
<p class="public-title011">
	<%=oblog.logined_uname%>����Ŀ��������
</p>
	<div id="toplogo"></div>
	<div id="sub">
	<img src='/images/[MyHomestay]Manager.gif' /><a href="HomestayBackctrl-index.asp" target="_top">������ҳ</a>| <img src='/images/logout.gif' /><a href="HomestayBackctrl-index.asp?t=logout&re=1">�˳���̨</a>
	
	
	<br />
	<br /><%=uurl%>|<img src='/images/up.gif' /><a href="user_update.asp">����վ��</a>
	</div>
	<div class="menu-t1" onClick="menu(oblog_1)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> ��Ŀ���¹���</div>
	<ul id="oblog_1">
	<li><a href="user_post.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/11.gif" align="absmiddle" /> �������</a></li>
	<li><a href="user_blogmanage.asp?usersearch=5"><img src="oBlogStyle/OblogStyleAdminImages/li/12.gif" align="absmiddle" /> �ݸ���</a></li>
	<li><a href="user_blogmanage.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/13.gif" align="absmiddle" /> ��������</a></li>
	<!--<li><a href="user_blogmanage.asp?action=downlog"><img src="oBlogStyle/OblogStyleAdminImages/li/14.gif" align="absmiddle" /> ��������</a></li>-->
	<!--<li><a href="user_subject.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/16.gif" align="absmiddle" /> ���·������</a></li>-->
	</ul>
	
	<!--<div class="menu-t1" onClick="menu(oblog_3)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> ��������</div>
	<ul id="oblog_3">
	<li><a href="user_comments.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/31.gif" align="absmiddle" /> ���۹���</a></li>
	<li><a href="user_messages.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/32.gif" align="absmiddle" /> ���Թ���</a></li>
	</ul>-->
	
	<div class="menu-t1" onClick="menu(oblog_4)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" />��Ŀģ������</div>
	<ul id="oblog_4" style="display: none;">
	
	<li><a href="user_template.asp?action=showconfig"><img src="oBlogStyle/OblogStyleAdminImages/li/41.gif" align="absmiddle" /> Ϊ��Ŀѡ��ģ��</a></li>
	<li><a href="user_template.asp?action=Add_user_skin_main_Nav"><img src="oBlogStyle/OblogStyleAdminImages/li/41.gif" align="absmiddle" /> �޸���Ŀ����ģ��</a></li>
	<li><a href="user_template.asp?action=modiconfig&amp;editm=1"><img src="oBlogStyle/OblogStyleAdminImages/li/42.gif" align="absmiddle" /> �޸�������ģ��</a></li>
	<li><a href="user_template.asp?action=modiviceconfig&amp;editm=1"><img src="oBlogStyle/OblogStyleAdminImages/li/43.gif" align="absmiddle" /> �޸����ݸ�ģ��</a></li>
	<li><a href="user_template.asp?action=bakskin"><img src="oBlogStyle/OblogStyleAdminImages/li/44.gif" align="absmiddle" /> ����ģ��</a></li>
	<!--<li><a href="user_template.asp?action=good"><img src="oBlogStyle/OblogStyleAdminImages/li/45.gif" align="absmiddle" /> �Ƽ��ҵ�ģ��</a></li>-->
	</ul>
	<%If EN_PHOTO=1 Then%>
	<div class="menu-t1" onClick="menu(oblog_2)" style="display:none;"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> ������</div>
	<ul id="oblog_2" style="display: none;">
	<li><a href="user_post.asp?t=1"><img src="oBlogStyle/OblogStyleAdminImages/li/21.gif" align="absmiddle" /> ���ͼƬ</a></li>
	<li><a href="user_blogmanage.asp?t=1"><img src="oBlogStyle/OblogStyleAdminImages/li/22.gif" align="absmiddle" /> ����ͼƬ</a></li>
	<!--<li><a href="user_subject.asp?t=1"><img src="oBlogStyle/OblogStyleAdminImages/li/24.gif" align="absmiddle" /> ���������</a></li>-->
	</li>
	</ul>
	<%end if%>
	<div class="menu-t1" onClick="menu(oblog_7)" style="display:none;"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" />��Ŀ�ļ�����</div>
	<ul id="oblog_7" style="display:none ;">
	<li><a href="user_files.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/71.gif" align="absmiddle" />��Ŀ�õ��������ļ�</a></li>
	<li><a href="user_files.asp?usersearch=1"><img src="oBlogStyle/OblogStyleAdminImages/li/72.gif" align="absmiddle" />��Ŀ�õ���ͼƬ�ļ�</a></li>
	<li><a href="user_files.asp?usersearch=2"><img src="oBlogStyle/OblogStyleAdminImages/li/73.gif" align="absmiddle" />��Ŀ�õ���ѹ���ļ�</a></li>
	<li><a href="user_files.asp?usersearch=3"><img src="oBlogStyle/OblogStyleAdminImages/li/74.gif" align="absmiddle" />��Ŀ�õ����ĵ��ļ�</a></li>
	<li style="display:"><a href="user_files.asp?usersearch=4"><img src="oBlogStyle/OblogStyleAdminImages/li/75.gif" align="absmiddle" />��Ŀ�õ���ý���ļ�</a></li>
	</ul>
	
	<!--<div class="menu-t1" onClick="menu(oblog_6)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> �ŶӺ���</div>
	<ul id="oblog_6" style="display: none;">
	<li><a href="user_friends.asp?action=add"><img src="oBlogStyle/OblogStyleAdminImages/li/61.gif" align="absmiddle" /> ��Ӻ���</a></li>
	<li><a href="user_friends.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/62.gif" align="absmiddle" /> ���ѹ���</a></li>
	<li><a href="user_friends.asp?action=add&amp;type=black"><img src="oBlogStyle/OblogStyleAdminImages/li/63.gif" align="absmiddle" /> ��Ӻ�����</a></li>
	<li><a href="user_friends.asp?usersearch=1"><img src="oBlogStyle/OblogStyleAdminImages/li/64.gif" align="absmiddle" /> ���������</a></li>
	<li><a href="user_team.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/65.gif" align="absmiddle" /> �û��Ŷ�</a></li>
	</ul>-->
	
	
	<div class="menu-t1" onClick="menu(oblog_5)"><img src="oBlogStyle/OblogStyleAdminImages/li/t1.gif" /> ��������</div>
	<ul id="oblog_5" style="display:none ;">
	<li><a href="user_setting.asp"><img src="oBlogStyle/OblogStyleAdminImages/li/51.gif" align="absmiddle" /> ��Ŀ����</a></li>
	<li><a href="user_setting.asp?action=placard"><img src="oBlogStyle/OblogStyleAdminImages/li/52.gif" align="absmiddle" /> �޸���Ŀ����</a></li>
	<li><a href="user_setting.asp?action=links"><img src="oBlogStyle/OblogStyleAdminImages/li/53.gif" align="absmiddle" /> ��Ŀ������������</a></li>
	<li><a href="user_setting.asp?action=userinfo"><img src="oBlogStyle/OblogStyleAdminImages/li/54.gif" align="absmiddle" /> �޸���Ŀ����</a></li>
	<li><a href="user_setting.asp?action=userpassword"><img src="oBlogStyle/OblogStyleAdminImages/li/55.gif" align="absmiddle" /> �޸���Ŀ��¼����</a></li>
	<li><a href="user_setting.asp?action=blogpassword"><img src="oBlogStyle/OblogStyleAdminImages/li/57.gif" align="absmiddle" /> ����������Ŀ</a></li>
	<!--<li><a href="user_setting.asp?action=blogstar"><img src="oBlogStyle/OblogStyleAdminImages/li/58.gif" align="absmiddle" /> �����û�����</a></li>-->
	</ul>
