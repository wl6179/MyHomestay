<!--#include file="inc/inc_sys.asp"-->
<%
dim action
action=trim(request("action"))
if action="saveconfig" then
	call saveconfig()
else
	call showconfig()
end if

sub showconfig()
dim rs
set rs=oblog.execute("select * from oblog_setup")
%>
<html>
<head>
<title>վ������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor" >
  <br>
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" Class="border">
    <tr class="topbg">       
    <td height="22" colspan=2 align=center><b>�� վ �� ��</b></td>
    </tr>
    <tr > 
      <td width="70" height="30"><strong>��������</strong></td>      
    <td height="30"><a href="admin_setup.asp#SiteInfo">��վ��Ϣ����</a> | <a href="admin_setup.asp#SiteOption">��վѡ������</a> 
      | <a href="#user">�û���Ϣѡ��</a> | <a href="#show">������ʾѡ��</a> | <a href="#upload">�ϴ�ѡ��</a></td>
    </tr>
</table>

<form method="POST" action="admin_setup.asp" id="form1" name="form1">
  <table width="98%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" Class="border">
    <tr > 
      <td height="22" colspan="2" class="topbg"> <a name="SiteInfo"></a><strong>��վ��Ϣ����</strong></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="348" height="25" >��վ���ƣ����50���ֽڣ���</td>
      <td width="409" height="25" > <input name="site_name" type="text" id="site_name" value="<%=oblog.filt_html(rs("site_name"))%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="348" height="25" >��վ���⣨���50���ֽڣ���</td>
      <td width="409" height="25" > <input name="site_title" type="text" id="site_title" value="<%=oblog.filt_html(rs("site_title"))%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="348" height="25" >��վ��ַ��<br>
        ��Ҫ������д����URL��ַ,��http://www.myhomestay.com.cn/blog/,<font color="#FF0000">����ʡ������/��</font>,�����ý�Ӱ�쵽rss��trackback���������С�</td>
      <td width="409" height="25" > <input name="site_path" type="text" id="site_path" value="<%=rs("site_path")%>" size="40" maxlength="255"> 
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="348" height="25" > �������������밴��MyHomestay.com.cn��������ʽ��д�����ж����������������&quot;|&quot;������<font color="#FF0000">��رն���������������</font>��</td>
      <td width="409" height="25" > <input name="site_domain" type="text" id="site_domain" value="<%=rs("site_domain")%>" size="40" maxlength="255"> 
      </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >վ��ؼ��֣������ױ����������ҵ�,&quot;,&quot;�Ÿ�����</td>
      <td height="25" ><input name="site_keyword" type="text" id="site_keyword" value="<%=rs("site_keyword")%>" size="40" maxlength="100"></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�Ƿ������������û����ӣ�</td>
      <td height="25" ><input type="radio" name="is_secondary_domain" value=1 <%if rs("is_secondary_domain")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_secondary_domain" value=0 <%if rs("is_secondary_domain")=0  then response.write "checked"%>>
        ��&nbsp;<font color="#FF0000">(��رջ�֧�ֶ�����������ѡ���)</font></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td height="25" >վ���Ȩ��Ϣ����ʾ��ϵͳҳ��ײ�����</td>
      <td height="25" ><textarea name="site_copyright" cols="55" rows="5" id="site_copyright"><%=rs("site_copyright")%></textarea> 
        <br> <a href="javascript:chang_size(-3,'site_copyright')"><img src="images/minus.gif" unselectable="on" border='0'></a> 
        <a href="javascript:chang_size(3,'site_copyright')"><img src="images/plus.gif" unselectable="on" border='0'></a> 
      </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="348" height="25" >վ�����䣺</td>
      <td width="409" height="25" > <input name="master_email" type="text" id="master_email" value="<%=rs("master_email")%>" size="40" maxlength="100"> 
      </td>
    </tr>
    <tr> 
      <td height="25" colspan="2" class="topbg"><a name="SiteOption"></a><b>��վѡ������</b></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="348" height="25" >�Ƿ���������ˣ�������˺������ʾ����</td>
      <td height="25" > <input type="radio" name="is_log_check" value=1 <%if rs("is_log_check")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_log_check" value=0 <%if rs("is_log_check")=0  then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�����ļ���Ĭ�����ɸ�ʽ��</td>
      <td height="25" > <input type="radio" name="df_filename" value=0 <%if rs("df_filename")=0 then response.write "checked"%>>
        ����ID�Զ���� &nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="radio" name="df_filename" value=1 <%if rs("df_filename")=1  then response.write "checked"%>>
        ���·���ʱ�� </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�������£�ϵͳ���·����Ƿ�Ϊ���룺</td>
      <td height="25" ><input type="radio" name="is_need_classid" value=1 <%if rs("is_need_classid")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_need_classid" value=0 <%if rs("is_need_classid")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�Ƿ���������ȫ������(����ر�)��</td>
      <td height="25" > <input type="radio" name="is_searchlogtext" value=1 <%if rs("is_searchlogtext")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_searchlogtext" value=0 <%if rs("is_searchlogtext")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >ÿ�������û�����ȫվ���ٴΣ�</td>
      <td height="25" ><input name="oneday_update" type="text" id="oneday_update" value="<%=rs("oneday_update")%>" size="10" maxlength="10">
        ��(����Ϊ0�򲻽�������)</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td height="25" >�Ƿ������οͷ������ۼ����ԣ�</td>
      <td height="25" > <input type="radio" name="is_enguest_comment" value=1 <%if rs("is_enguest_comment")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_enguest_comment" value=0 <%if rs("is_enguest_comment")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >��������ʱ���������ֳ��ֶ����֣������Զ�תΪ����ˣ�</td>
      <td height="25" ><input name="max_badstr_num" type="text" id="max_badstr_num" value="<%=rs("max_badstr_num")%>" size="10" maxlength="10">
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td height="25" >�û�ע���Ƿ���Ҫ��֤�룺</td>
      <td height="25" ><input type="radio" name="is_code_reg" value=1 <%if rs("is_code_reg")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_code_reg" value=0 <%if rs("is_code_reg")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td height="25" >�û���½�Ƿ���Ҫ��֤�룺</td>
      <td height="25" ><input type="radio" name="is_code_login" value=1 <%if rs("is_code_login")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_code_login" value=0 <%if rs("is_code_login")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td height="25" >�û��������ۣ������Ƿ���Ҫ��֤�룺</td>
      <td height="25" ><input type="radio" name="is_code_comment" value=1 <%if rs("is_code_comment")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_code_comment" value=0 <%if rs("is_code_comment")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <!--- <tr class="tdbg" onmouseout="this.style.backgroundColor=''" onmouseover="this.style.backgroundColor='#BFDFFF'"> 
      <td height="25" >�û���������Ϣ�Ƿ���Ҫ��֤�룺</td>
      <td height="25" ><input type="radio" name="is_code_pm" value=1 <%'if rs("is_code_pm")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_code_pm" value=0 <%'if rs("is_code_pm")=0 then response.write "checked"%>>
        ��</td>
		
    </tr>-->
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td height="25" >�û����������Ƿ���Ҫ��֤�룺</td>
      <td height="25" ><input type="radio" name="is_code_addblog" value=1 <%if rs("is_code_addblog")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_code_addblog" value=0 <%if rs("is_code_addblog")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�û�ҳ��ͳ�Ʒ�ˢ��ʱ�䣺</td>
      <td height="25" ><input name="dis_refutime" type="text" id="dis_refutime" value="<%=rs("dis_refutime")%>" size="10" maxlength="10">
        �� </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�������£����ԣ����۵�ʱ������</td>
      <td height="25" ><input name="add_pitchtime" type="text" id="add_pitchtime" value="<%=rs("add_pitchtime")%>" size="10" maxlength="10">
        �� </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td height="25" >��ҳ��̬�ļ��ĸ���ʱ�䣨��������300�����ϣ����򽫼��ķѷ�������Դ����</td>
      <td height="25" ><input name="time_index_update" type="text" id="time_index_update" value="<%=rs("time_index_update")%>" size="10" maxlength="10">
        �� </td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td height="25" >�û�������Ϣ������</td>
      <td height="25" ><input name="pm_num" type="text" id="pm_num" value="<%=rs("pm_num")%>" size="10" maxlength="10">
        �� </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >��ƪ�����������������</td>
      <td height="25" ><input name="max_log_word" type="text" id="max_log_word" value="<%=rs("max_log_word")%>" size="10" maxlength="10">
        ��(Ӣ���ַ�) </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�ظ��������������������</td>
      <td height="25" ><input name="max_comment_word" type="text" id="max_comment_word" value="<%=rs("max_comment_word")%>" size="10" maxlength="10">
        ��(Ӣ���ַ�) </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >ת���û���ҳ�Ƿ�ת��index�ļ���</td>
      <td height="25" ><input type="radio" name="index_to_file" value=1 <%if rs("index_to_file")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="index_to_file" value=0 <%if rs("index_to_file")=0  then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�����û��޸Ķ���������</td>
      <td height="25" ><input type="radio" name="enchangdomain" value=1 <%if rs("enchangdomain")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="enchangdomain" value=0 <%if rs("enchangdomain")=0  then response.write "checked"%>>
        ��</td>
    </tr>
    <tr> 
      <td height="25" colspan="2" class="topbg"><a name="show"></a><b>������ʾѡ��</b></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >ϵͳ�����б�ÿҳ��ʾ���������±���</td>
      <td height="25" ><input name="show_log_num" type="text" id="show_log_num" value="<%=rs("show_log_num")%>" size="10" maxlength="10"> 
      </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" ><p>ϵͳ�����б����ǰ����ƪ����</p></td>
      <td height="25" > <input name="get_log_num" type="text" id="get_log_num" value="<%=rs("get_log_num")%>" size="10" maxlength="10"></td>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >ͼƬ�б�ÿҳ��ʾ���ٸ�ͼƬ</td>
      <td height="25" ><input name="show_photonum" type="text" id="show_photonum" value="<%=rs("show_photonum")%>" size="10" maxlength="10"> 
      </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" ><p>ͼƬ�б����ǰ���ٸ�ͼƬ</p></td>
      <td height="25" > <input name="get_photonum" type="text" id="get_photonum" value="<%=rs("get_photonum")%>" size="10" maxlength="10"></td>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�����Ƿ�˳�����У�</td>
      <td height="25" ><input type="radio" name="show_comment_asc" value=1 <%if rs("show_comment_asc")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="show_comment_asc" value=0 <%if rs("show_comment_asc")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >ҳ����������ʱ��ʾ�ַ�</td>
      <td height="25" ><input name="loadingstr" type="text" id="loadingstr" value="<%=rs("loadingstr")%>" size="30" maxlength="50"></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td height="25" >�û��б�ÿҳ��ʾ����</td>
      <td height="25" ><input name="show_listuser_num" type="text" id="show_listuser_num" value="<%=rs("show_listuser_num")%>" size="10" maxlength="10"></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td height="25" >ͼƬ�Զ���С��ȣ�Ϊ�㲻���ţ�</td>
      <td height="25" ><input name="show_imgw_num" type="text" id="show_imgw_num" value="<%=rs("show_imgw_num")%>" size="10" maxlength="10">
        ����</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="348" height="25" >ͼƬ�Ƿ������������ţ�</td>
      <td height="25" > <input type="radio" name="show_img_mouse" value=1 <%if rs("show_img_mouse")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="show_img_mouse" value=0 <%if rs("show_img_mouse")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="348" height="25" >������ʾ�����Ƿ�ʹ��htm���ǿ�����ˣ�<br>
        ����ѡ�������г�ͼƬ����ı�Ƕ��������˵���</td>
      <td height="25" > <input type="radio" name="is_log_profilt" value=1 <%if rs("is_log_profilt")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_log_profilt" value=0 <%if rs("is_log_profilt")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr> 
      <td height="25" colspan="2" class="topbg"><a name="user" id="user"></a><strong>�û�ѡ��</strong></td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="348" height="25" >�Ƿ��������û�ע�᣺</td>
      <td height="25" > <input type="radio" name="is_enreg" value=1 <%if rs("is_enreg")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_enreg" value=0 <%if rs("is_enreg")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#BFDFFF'"> 
      <td width="348" height="25" > <p>���û�ע���Ƿ���Ҫ����Ա��֤��<br>
          ��ѡ���ǣ����û�������ͨ������Ա��֤����������ɹ���ʽע���û���</p></td>
      <td height="25" > <input type="radio" name="is_admin_chkreg" value=1 <%if rs("is_admin_chkreg")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_admin_chkreg" value=0 <%if rs("is_admin_chkreg")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <!--
    <tr class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''"> 
      <td height="25" >VIP�û��Ƿ��ܲ鿴�������¼��������£�</td>
      <td height="25" ><input name="is_vip_prosee" type="radio" value=1 <%'if rs("is_vip_prosee")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="is_vip_prosee" value=0 <%'if rs("is_vip_prosee")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr> -->
    <td height="25" colspan="2" class="topbg"><a name="upload" id="user"></a><strong>�ϴ�ѡ��</strong></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�Ƿ�������ͨ��Ա�ϴ��ļ���</td>
      <td height="25" ><input type="radio" name="upfile_user_en" value=1 <%if rs("upfile_user_en")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="upfile_user_en" value=0 <%if rs("upfile_user_en")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >��ͨ��Ա�ϴ��ļ����ͣ�</td>
      <td height="25" ><input name="upfile_user_type" type="text" id="upfile_user_type" value="<%=rs("upfile_user_type")%>" size="60" maxlength="200"></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >��ͨ��Ա�ϴ������ļ���С��</td>
      <td height="25" ><input name="upfile_user_onesize" type="text" id="upfile_user_onesize" value="<%=rs("upfile_user_onesize")%>" size="10" maxlength="10">
        KB</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�Ƿ�����VIP��Ա�ϴ��ļ���</td>
      <td height="25" ><input type="radio" name="upfile_vip_en" value=1 <%if rs("upfile_vip_en")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="upfile_vip_en" value=0 <%if rs("upfile_vip_en")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >VIP��Ա�ϴ��ļ����ͣ�</td>
      <td height="25" ><input name="upfile_vip_type" type="text" id="upfile_vip_type" value="<%=rs("upfile_vip_type")%>" size="60" maxlength="200"></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >VIP��Ա�ϴ������ļ���С��</td>
      <td height="25" ><input name="upfile_vip_onesize" type="text" id="upfile_vip_onesize" value="<%=rs("upfile_vip_onesize")%>" size="10" maxlength="10">
        KB</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�Ƿ�����ǰ̨����Ա�ϴ��ļ���</td>
      <td height="25" ><input type="radio" name="upfile_admin_en" value=1 <%if rs("upfile_admin_en")=1 then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="upfile_admin_en" value=0 <%if rs("upfile_admin_en")=0 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >ǰ̨����Ա�ϴ��ļ����ͣ�</td>
      <td height="25" ><input name="upfile_admin_type" type="text" id="upfile_admin_type" value="<%=rs("upfile_admin_type")%>" size="60" maxlength="200"></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >ǰ̨����Ա�ϴ������ļ���С��</td>
      <td height="25" ><input name="upfile_admin_onesize" type="text" id="upfile_admin_onesize" value="<%=rs("upfile_admin_onesize")%>" size="10" maxlength="10">
        KB</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >��ͨ��Ա�ϴ��ռ��С��</td>
      <td height="25" ><input name="upfile_user_maxsize" type="text" id="upfile_user_maxsize" value="<%=rs("upfile_user_maxsize")%>" size="10" maxlength="10">
        KB</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >VIP��Ա�ϴ��ռ��С��</td>
      <td height="25" ><input name="upfile_vip_maxsize" type="text" id="upfile_vip_maxsize" value="<%=rs("upfile_vip_maxsize")%>" size="10" maxlength="10">
        KB</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >ǰ̨����Ա�ϴ��ռ��С��</td>
      <td height="25" ><input name="upfile_admin_maxsize" type="text" id="upfile_admin_maxsize" value="<%=rs("upfile_admin_maxsize")%>" size="10" maxlength="10">
        KB</td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" ><p>ѡȡ�ϴ������<br>
          (���Ե�<a href="http://www.myhomestay.com.cn" target="_blank">http://www.myhomestay.com.cn</a>����Aspupload3.0���) 
        </p></td>
      <td height="25" ><select name="upset_uptype" id="upset_uptype" onChange="chkselect(options[selectedIndex].value,'know2');">
          <option value="999">�ر� 
          <option value="0">������ϴ��� 
          <option value="1">Aspupload3.0��� 
          <option value="2">SA-FileUp 4.0��� </select> <div id="know2"></div></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >ͼƬ����ͼ��ˮӡ���ÿ��أ�<br>
        (�������谲װaspjepg������ɵ�<a href="http://www.myhomestay.com.cn" target="_blank">http://www.myhomestay.com.cn</a>����) 
      </td>
      <td height="25" ><SELECT name="upset_isDraw" id="upset_isDraw">
          <OPTION value="0" >�ر�����ͼ��ˮӡЧ��</OPTION>
          <OPTION value="1">��������ͼ��ˮӡ����Ч��(�Ƽ�)</OPTION>
          <OPTION value="2">��������ͼ��ˮӡͼƬЧ��</OPTION>
        </SELECT></br>
        <%If IsObjInstalled("Persits.Jpeg") Then Response.Write "aspjepg���<font color=red><b>��</b>������֧��!</font>" Else Response.Write "aspjepg���<b>��</b>��������֧��!" %> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�ϴ�ͼƬ���ˮӡ������Ϣ����Ϊ�ջ�0��:</td>
      <td height="25" ><INPUT TYPE="text" NAME="upset_Drawtext" size=40 value="<%=rs("upset_Drawtext")%>"></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�ϴ����ˮӡ�����С:</td>
      <td height="25" > <INPUT TYPE="text" NAME="upset_Drawfontsize" size=10 value="<%=rs("upset_Drawfontsize")%>"> 
        <b>px</b> </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�ϴ����ˮӡ������ɫ:</td>
      <td height="25" ><INPUT TYPE="text" NAME="upset_Drawcolor" ID="d_upset_Drawcolor" size=10 value="<%if rs("upset_Drawcolor")="" then response.Write("#FFFFFF") else response.Write(rs("upset_Drawcolor"))%>"> 
        <img border=0 id="s_upset_Drawcolor" src="images/rect.gif" style="cursor:pointer;background-Color:<%=rs("upset_Drawcolor")%>;" onClick="SelectColor('upset_Drawcolor');" title="ѡȡ��ɫ!"> 
      </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�ϴ����ˮӡ��������:</td>
      <td height="25" ><SELECT name="upset_Drawfont" id="upset_Drawfont">
          <option value="����">����</option>
          <option value="����_GB2312">����</option>
          <option value="������">������</option>
          <option value="����">����</option>
          <option value="����">����</option>
          <OPTION value="Andale Mono" selected>Andale Mono</OPTION>
          <OPTION value=Arial>Arial</OPTION>
          <OPTION value="Arial Black">Arial Black</OPTION>
          <OPTION value="Book Antiqua">Book Antiqua</OPTION>
          <OPTION value="Century Gothic">Century Gothic</OPTION>
          <OPTION value="Comic Sans MS">Comic Sans MS</OPTION>
          <OPTION value="Courier New">Courier New</OPTION>
          <OPTION value=Georgia>Georgia</OPTION>
          <OPTION value=Impact>Impact</OPTION>
          <OPTION value=Tahoma>Tahoma</OPTION>
          <OPTION value="Times New Roman" >Times New Roman</OPTION>
          <OPTION value="Trebuchet MS">Trebuchet MS</OPTION>
          <OPTION value="Script MT Bold">Script MT Bold</OPTION>
          <OPTION value=Stencil>Stencil</OPTION>
          <OPTION value=Verdana>Verdana</OPTION>
          <OPTION value="Lucida Console">Lucida Console</OPTION>
        </SELECT></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�ϴ�ˮӡ�����Ƿ����:</td>
      <td height="25" ><SELECT name="upset_DrawFontBold" id="upset_DrawFontBold">
          <OPTION value=0>��</OPTION>
          <OPTION value=1>��</OPTION>
        </SELECT></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�ϴ�ͼƬ���ˮӡLOGOͼƬ��Ϣ����Ϊ�ջ�0��:<br>
        ��дLOGO��ͼƬ���·��</td>
      <td height="25" ><INPUT TYPE="text" NAME="upset_Drawpic" size=40 value="<%=rs("upset_Drawpic")%>"></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�ϴ�ͼƬ���ˮӡ͸����:</td>
      <td height="25" ><INPUT TYPE="text" NAME="upset_DrawGraph" size=10 value="<%=rs("upset_DrawGraph")%>">
        ��60%����д0.6 </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >ˮӡͼƬȥ����ɫ:<br>
        ����Ϊ����ˮӡͼƬ��ȥ����ɫ��</td>
      <td height="25" ><INPUT TYPE="text" NAME="upset_Drawpiccolor" ID="d_upset_Drawpiccolor" size=10 value="<%=rs("upset_Drawpiccolor")%>"> 
        <img border=0 id="s_upset_Drawpiccolor" src="images/rect.gif" style="cursor:pointer;background-Color:<%=rs("upset_Drawpiccolor")%>;" onClick="SelectColor('upset_Drawpiccolor');" title="ѡȡ��ɫ!"> 
      </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >ˮӡ���ֻ�ͼƬ�ĳ���������:<br>
        ��ˮӡͼƬ�Ŀ�Ⱥ͸߶ȡ�</td>
      <td height="25" >��ȣ� 
        <INPUT TYPE="text" NAME="upset_DrawWidth" size=10 value="<%=rs("upset_DrawWidth")%>">
        ���� �߶ȣ� 
        <INPUT TYPE="text" NAME="upset_DrawHight" size=10 value="<%=rs("upset_DrawHight")%>">
        ���� </td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >�ϴ�ͼƬ���ˮӡLOGOλ������ :</td>
      <td height="25" ><SELECT NAME="upset_DrawXYType" id="upset_DrawXYType">
          <option value="0" <%if rs("upset_DrawXYType")=0 then response.Write("selected")%>>����</option>
          <option value="1" <%if rs("upset_DrawXYType")=1 then response.Write("selected")%>>����</option>
          <option value="2" <%if rs("upset_DrawXYType")=2 then response.Write("selected")%>>����</option>
          <option value="3" <%if rs("upset_DrawXYType")=3 then response.Write("selected")%>>����</option>
          <option value="4" <%if rs("upset_DrawXYType")=4 then response.Write("selected")%>>����</option>
        </SELECT></td>
    </tr>
    <tr class="tdbg" onMouseOver="this.style.backgroundColor='#BFDFFF'" onMouseOut="this.style.backgroundColor=''"> 
      <td height="25" >&nbsp;</td>
      <td height="25" >&nbsp; </td>
    </tr>
    <tr> 
      <td height="40" colspan="2" align="center" class="tdbg" > <input name="Action" type="hidden" id="Action" value="saveconfig"> 
        <input name="cmdSave" type="submit" id="cmdSave" value=" �������� " > </td>
    </tr>
  </table>

</form>
<SCRIPT LANGUAGE="JavaScript">
CheckSel('upset_uptype','<%=rs("upset_uptype")%>');
CheckSel('upset_isDraw','<%=rs("upset_isDraw")%>');
CheckSel('upset_Drawfont','<%=rs("upset_Drawfont")%>');
CheckSel('upset_DrawFontBold','<%=rs("upset_DrawFontBold")%>');
CheckSel('upset_DrawXYType','<%=rs("upset_DrawXYType")%>');

function SelectColor(what){
	var dEL = document.all("d_"+what);
	var sEL = document.all("s_"+what);
	var arr = showModalDialog("../images/selcolor.html", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
	if (arr) {
		dEL.value=arr;
		sEL.style.backgroundColor=arr;
	}
}
function chkselect(s,divid)
{
var divname='Issubport';
var chkreport;
	s=Number(s)
	if (divid=="know1")
	{
	divname=divname+s;
	}
	if (divid=="know2")
	{
	s+=4;
	if (s==1003){s=999;}
	divname=divname+s;
	}
	if (divid=="know3")
	{
	s+=9;
	if (s==1008){s=999;}
	divname=divname+s;
	}
document.getElementById(divid).innerHTML=divname;
chkreport=document.getElementById(divname).innerHTML;
document.getElementById(divid).innerHTML=chkreport;
}

</script>
<div id="Issubport0" style="display:none">��ѡ��EMAIL�����</div>
<%
Dim InstalledObjects(12),i
InstalledObjects(1) = "JMail.Message"				'JMail 4.3
InstalledObjects(2) = "CDONTS.NewMail"				'CDONTS
InstalledObjects(3) = "Persits.MailSender"			'ASPEMAIL
'-----------------------
InstalledObjects(4) = "Adodb.Stream"				'Adodb.Stream
InstalledObjects(5) = "Persits.Upload"				'Aspupload3.0
InstalledObjects(6) = "SoftArtisans.FileUp"			'SA-FileUp 4.0
InstalledObjects(7) = "DvFile.Upload"				'DvFile-Up V1.0
'-----------------------
InstalledObjects(9) = "CreatePreviewImage.cGvbox"	'CreatePreviewImage
InstalledObjects(10) = "Persits.Jpeg"				'AspJpeg
InstalledObjects(11) = "Persits.Jpeg"		'SoftArtisans ImgWriter V1.21
InstalledObjects(12) = "sjCatSoft.Thumbnail"		'sjCatSoft.Thumbnail V2.6

For i=1 to 12
	Response.Write "<div id=""Issubport"&i&""" style=""display:none"">"
	If IsObjInstalled(InstalledObjects(i)) Then Response.Write InstalledObjects(i)&":<font color=red><b>��</b>������֧��!</font>" Else Response.Write InstalledObjects(i)&"<b>��</b>��������֧��!" 
	Response.Write "</div>"
Next
%>
</body>
</html>
<%
set rs=nothing
end sub

sub saveconfig()
dim rs,sql
if not IsObject(conn) then link_database
set rs=server.CreateObject("adodb.recordset")
sql="select * from oblog_setup"
rs.open sql,conn,1,3
rs("site_name")=trim(request("site_name"))
rs("site_title")=trim(request("site_title"))
rs("site_path")=trim(request("site_path"))
rs("site_domain")=trim(request("site_domain"))
rs("site_copyright")=trim(request("site_copyright"))
rs("master_email")=trim(request("master_email"))
rs("show_log_num")=trim(request("show_log_num"))
rs("is_log_check")=trim(request("is_log_check"))
rs("is_comment_check")=0'trim(request("commentcheck"))
rs("is_secondary_domain")=trim(request("is_secondary_domain"))
'rs("timediff")=trim(request("timediff"))
'rs("sessiontimeout")=trim(request("sessiontimeout"))
rs("is_enreg")=trim(request("is_enreg"))
rs("is_admin_chkreg")=trim(request("is_admin_chkreg"))
rs("is_enguest_comment")=trim(request("is_enguest_comment"))
rs("site_keyword")=trim(request("site_keyword"))
rs("show_listuser_num")=clng(trim(request("show_listuser_num")))
if request("get_log_num")="" then
	rs("get_log_num")=500
else
	rs("get_log_num")=clng(trim(request("get_log_num")))
end if

if request("show_imgw_num")="" then
	rs("show_imgw_num")=0
else
	rs("show_imgw_num")=clng(trim(request("show_imgw_num")))
end if
rs("show_img_mouse")=trim(request("show_img_mouse"))
rs("is_log_profilt")=trim(request("is_log_profilt"))
'rs("is_vip_prosee")=trim(request("is_vip_prosee"))

rs("upfile_user_en")=trim(request("upfile_user_en"))
rs("upfile_user_type")=trim(request("upfile_user_type"))
if request("upfile_user_onesize")="" then
	rs("upfile_user_onesize")=0
else
	rs("upfile_user_onesize")=clng(trim(request("upfile_user_onesize")))
end if

rs("upfile_vip_en")=trim(request("upfile_vip_en"))
rs("upfile_vip_type")=trim(request("upfile_vip_type"))
if request("upfile_vip_onesize")="" then
	rs("upfile_vip_onesize")=0
else
	rs("upfile_vip_onesize")=clng(trim(request("upfile_vip_onesize")))
end if
rs("upfile_admin_en")=trim(request("upfile_admin_en"))
rs("upfile_admin_type")=trim(request("upfile_admin_type"))
if request("upfile_admin_onesize")="" then
	rs("upfile_admin_onesize")=0
else
	rs("upfile_admin_onesize")=clng(trim(request("upfile_admin_onesize")))
end if

if request("max_badstr_num")="" then
	rs("max_badstr_num")=0
else
	rs("max_badstr_num")=clng(trim(request("max_badstr_num")))
end if

rs("show_comment_asc")=trim(request("show_comment_asc"))
'response.Write clng(trim(request("reg_userlevel"))
if request("upfile_user_maxsize")="" then
	rs("upfile_user_maxsize")=0
else
	rs("upfile_user_maxsize")=clng(trim(request("upfile_user_maxsize")))
end if
if request("upfile_vip_maxsize")="" then
	rs("upfile_vip_maxsize")=0
else
	rs("upfile_vip_maxsize")=clng(trim(request("upfile_vip_maxsize")))
end if
if request("upfile_admin_maxsize")="" then
	rs("upfile_admin_maxsize")=0
else
	rs("upfile_admin_maxsize")=clng(trim(request("upfile_admin_maxsize")))
end if

rs("is_need_classid")=trim(request("is_need_classid"))
if request("add_pitchtime")="" then
	rs("add_pitchtime")=0
else
	rs("add_pitchtime")=clng(trim(request("add_pitchtime")))
end if
if request("time_index_update")="" then
	rs("time_index_update")=0
else
	rs("time_index_update")=clng(trim(request("time_index_update")))
end if
rs("is_code_reg")=cint(request("is_code_reg"))
rs("is_code_login")=cint(request("is_code_login"))
rs("is_code_comment")=cint(request("is_code_comment"))
'rs("is_code_pm")=cint(request("is_code_pm"))
rs("is_code_addblog")=cint(request("is_code_addblog"))
rs("pm_num")=cint(request("pm_num"))
rs("enchangdomain")=cint(request("enchangdomain"))
rs("dis_refutime")=CLNG(request("dis_refutime"))
rs("upset_uptype")=request("upset_uptype")
rs("upset_isDraw")=request("upset_isDraw")
rs("upset_Drawtext")=request("upset_Drawtext")
rs("upset_Drawfontsize")=request("upset_Drawfontsize")
rs("upset_Drawcolor")=request("upset_Drawcolor")
rs("upset_Drawfont")=request("upset_Drawfont")
rs("upset_DrawFontBold")=request("upset_DrawFontBold")
rs("upset_Drawpic")=request("upset_Drawpic")
if request("upset_DrawGraph")="" then
	rs("upset_DrawGraph")=0
else
	rs("upset_DrawGraph")=request("upset_DrawGraph")
end if
rs("upset_Drawpiccolor")=request("upset_Drawpiccolor")
if request("upset_DrawHight")="" then 
	rs("upset_DrawHight")=0
else
	rs("upset_DrawHight")=cint(request("upset_DrawHight"))
end if
if request("upset_DrawWidth")="" then
	rs("upset_DrawWidth")=0
else
	rs("upset_DrawWidth")=cint(request("upset_DrawWidth"))
end if
rs("upset_DrawXYType")=request("upset_DrawXYType")
rs("max_log_word")=request("max_log_word")
rs("max_comment_word")=request("max_comment_word")
rs("is_searchlogtext")=request("is_searchlogtext")
rs("get_photonum")=request("get_photonum")
rs("show_photonum")=request("show_photonum")
rs("oneday_update")=request("oneday_update")
If request("index_to_file")<>"" Then
	rs("index_to_file")=request("index_to_file")
Else
	rs("index_to_file")=0
End If
rs("loadingstr")=request("loadingstr")
If request("df_filename")<>"" then
	rs("df_filename")=request("df_filename")
Else
	rs("df_filename")=0
End If
rs.update
rs.close
set rs=nothing
oblog.reloadsetup
response.Redirect "admin_setup.asp"
end sub
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err = 0 Then IsObjInstalled = True
	If Err = -2147352567 Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

%>