<html>
<head>
<title>MyHomestay ��ǩ����˵��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">

<br>
<table width="90%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center"> 
    <td height=25 colspan="2" class="topbg"><strong>ϵͳģ����˵��</strong> 
  <tr> 
<td height=23 colspan="2" class="tdbg"><p><strong><font color=#0000ff>��ģ�棨��ϵͳ��ҳ��ֻ��index.asp�ļ���Ч����</font><br>
        ע��</strong>����ʹ����ȷ�ı�ǩ������������ܳ��ֲ���Ԥ֪�Ĵ���һ��Ϊ��ʾΪ�±�Խ�磩<br>
        ������ ��ɫ����Ϊ3.1�汾��������3.0�汾��ͬ�����ı�ǩ��</p></td>
  </tr>
  <tr> 
    <td width="25%" height=23 class="tdbg"><p>$show_log(����1,����2,����3,����4,����5,����6,����7,����8,����9)$<br>
      </p></td>
    <td width="75%" class="tdbg">�˱�ǵ������±����б����Ϣ������˵�����£�<br>
      ����1����������������<br>
      ����2�����±��ⳤ�ȣ����ַ�Ϊ��λ���������ֽ�������ʾ��<br>
      ����3�����򷽷���Ϊ1������ʱ�䣬Ϊ2���������Ϊ3���ظ�����<br>
      ����4���Ƿ񾫻���Ϊ1�����������£�Ϊ2���þ������¡�<br>
      ����5�����ö������ڵ����£�����Ϊ��λ��<br>
      ����6�����·���id��Ϊ0��������з�������¡�<br>
      ����7���Ƿ���ʾ����ϵͳ��������Ϊ1��ʾ��Ϊ0����ʾ��<br>
      ����8���Ƿ���ʾ����ר������Ϊ1��ʾ��Ϊ0����ʾ��<br>
      ����9����ʾ��Ϣ��1Ϊ��ʾ����ʱ����û���2Ϊ��ʾ����ʱ�䣬3Ϊ��ʾ�����û���4Ϊ��ʾ�����û��͵������5Ϊ��ʾ�������6Ϊ��ʾ�������ں��û���7Ϊ��ʾ�������ڣ�8Ϊ��ʾ�ظ�����0Ϊ����ʾ��</td>
  </tr>
  <tr> 
<td width="25%" height=23 class="tdbg"><p><font color="#FF0000">$show_userlog(����1,����2,����3,����4,����5,����6)$</font><br>
      </p></td>
    <td width="75%" class="tdbg"><font color="#FF0000">�˱�ǵ���ĳ�û����±����б����Ϣ������˵�����£�<br>
      	����1��userid��<br>
		����2����������������<br>
		����3�����±��ⳤ�ȣ����ַ�Ϊ��λ���������ֽ�������ʾ��<br>
		����4�����򷽷���Ϊ1������ʱ�䣬Ϊ2���������Ϊ3���ظ�����<br>
		����5���û�ר��id��Ϊ0����ø��û��������¡�<br>
		����6����ʾ��Ϣ��1Ϊ��ʾ�������ڣ�0Ϊ����ʾ��
	</font>
	</td>
  </tr>
  
  �˱�ǵ������±����б����Ϣ������˵�����£�


  <tr> 
    <td height=23 class="tdbg">$show_placard$</td>
    <td height=23 class="tdbg">�˱����ʾϵͳ���档</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_count$</td>
    <td height=23 class="tdbg"> �˱��վ��ͳ����Ϣ��</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_blogupdate(����1)$</td>
    <td height=23 class="tdbg"> �˱����ʾ���¸��������б�����1������������</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_userlogin$</td>
    <td height=23 class="tdbg"> �˱����ʾ��¼���ڡ� </td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_comment(����1,����2)$</td>
    <td height=23 class="tdbg"> �˱����ʾ���»ظ��б�����1����������������2���ظ����ⳤ�ȡ�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_subject(����1)$</td>
    <td height=23 class="tdbg"> �˱����ʾ�û�ר�������б�����1������������</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_bestblog(����1)$</td>
    <td height=23 class="tdbg">�˱����ʾ�Ƽ��û�������1������������</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_newblogger(����1)$</td>
    <td height=23 class="tdbg">�˱����ʾ����ע���û�������1������������</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_class(����1)$</td>
    <td height=23 class="tdbg"> �˱����ʾϵͳ�����б�����1��������ʾʱ��ÿ��������Ϊ0��������ʾ��</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_bloger(����1)$</td>
    <td height=23 class="tdbg"> �˱����ʾ�û��б�����1��������ʾʱ��ÿ��������Ϊ0��������ʾ��</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_friends$</td>
    <td height=23 class="tdbg"> �˱����ʾ�������ӡ�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_search(����1)$</td>
    <td height=23 class="tdbg"> �˱����ʾ������������1��Ϊ0������ʾ��Ϊ1��������ʾ��</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_cityblogger(����1)$</td>
    <td height=23 class="tdbg">�˱����ʾͬ���û�������������1��Ϊ0������ʾ��Ϊ1��������ʾ��<font color="#FF0000">�˱�ǩ�������ڸ�ģ�档</font> 
    </td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_blogstar$</td>
    <td height=23 class="tdbg">�˱�ǵ��������û����ǡ�</td>
  </tr>
    <tr> 
    <td height=23 class="tdbg">$show_blogstar2(����1,����2,����3,����4)$</td>
    <td height=23 class="tdbg">�˱�ǵ��������û����ǡ�����1��������Ŀ������2��ÿ����ʾ��Ŀ������3��ͼƬ��ȣ�����4��ͼƬ�߶ȡ�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_newphoto(����1,����2,����3,����4)$</td>
    <td height=23 class="tdbg">�˱�ǩ�������ͼƬ������1����������������2��0Ϊ������ʾ��1Ϊ������ʾ������3��ͼƬ��ȣ�����4��ͼƬ�߶ȡ�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_xml$</td>
    <td height=23 class="tdbg"> �˱����ʾrss���ӱ�־��</td>
  </tr>
  <tr> 
    <td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>��ģ�棨�Գ�index.asp�������ϵͳҳ����Ч����HomestayNav.asp,listblogger.asp�ļ��ȣ���</font></strong></td>
  </tr>
  <tr> 
    <td height=23 colspan="2" class="tdbg">����������ģ��ı�ǣ�������ͬ��</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_list$</td>
    <td height=23 class="tdbg"> ��Ҫ���˱����ʾ����ϵͳ��ҳ���������壬����ȥ������ֻ���ڸ�����ʹ�á�</td>
  </tr>
</table>
<br>
<table width="90%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center"> 
    <td height=25 colspan="2" class="topbg"><strong>�û�ģ����˵��</strong> 
  <tr> 
    <td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>��ģ�棺</font><br>
      </strong>��ģ��Ϊҳ������岿�֣�����css��ʽ���õȣ�������Dreamweave��Frontpage�б༭����ɺ󽫴���copy������ 
    
    </td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_log$</td>
    <td height=23 class="tdbg"> ��Ҫ���˱����ʾ�������岿�֣��������۵���Ϣ��</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_placard$ </td>
    <td height=23 class="tdbg">�˱����ʾ�û����档</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_calendar$</td>
    <td height=23 class="tdbg"> �˱����ʾ������</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_newblog$</td>
    <td height=23 class="tdbg"> �˱����ʾ���������б�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_comment$</td>
    <td height=23 class="tdbg"> �˱����ʾ���»ظ��б�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_subject$</td>
    <td height=23 class="tdbg"> �˱����ʾר����ࡣ</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_subject_l$</td>
    <td height=23 class="tdbg"> �˱�Ǻ�����ʾר����ࡣ </td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_newblog$</td>
    <td height=23 class="tdbg"> �˱����ʾ���������б�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_newmessage$ </td>
    <td height=23 class="tdbg">�˱����ʾ���������б�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_info$</td>
    <td height=23 class="tdbg"> �˱����ʾ��ʵ�����ƣ�ͳ����Ϣ�ȡ�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_login$</td>
    <td height=23 class="tdbg"> �˱����ʾ��¼���ڡ�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_links$</td>
    <td height=23 class="tdbg"> �˱����ʾ������Ϣ��</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_blogname$ </td>
    <td height=23 class="tdbg">�˱����ʾ�û���ʵ�����ƣ�������Ϊ������ʾ�û�id��</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_search$</td>
    <td height=23 class="tdbg"> �˱����ʾ��������</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_xml$</td>
    <td height=23 class="tdbg"> �˱����ʾrss���ӱ�־��</td>
  </tr>
  <tr> 
<td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>��ģ�棺</font><br>
      </strong>��ģ��Ϊ��ʾ�������ݲ��֡��������±��⣬����ʱ�䣬�������ݵ���Ϣ�İ������á�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_topic$</td>
    <td height=23 class="tdbg"> �˱����ʾ����ͼ�꣬ר���������±���,��</td>
  </tr>
  <tr> 
    <td width="18%" height=23 class="tdbg">$show_loginfo$</td>
    <td width="82%" class="tdbg">�˱����ʾ�������ߣ�����ʱ�����Ϣ��</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_logtext$</td>
    <td height=23 class="tdbg"> �˱����ʾ�������ġ�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_more$</td>
    <td height=23 class="tdbg"> �˱����ʾ�Ķ�ȫ��(����)���ظ�(����)���������ӡ�</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_emot$</td>
    <td height=23 class="tdbg">�˱�ǽ���ʾ��ʾ����ͼ�ꡣ</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_author$</td>
    <td height=23 class="tdbg">�˱�ǽ���ʾ��������</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_addtime$</td>
    <td height=23 class="tdbg">�˱�ǽ���ʾ����ʱ�䡣</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_topictxt$</td>
    <td height=23 class="tdbg">�˱�ǽ���ʾ���±��⡣</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_blogzhai$</td>
    <td height=23 class="tdbg">�˱����ʾ���뵽��ժ�����ӡ�</td>
  </tr>
</table>
</body>
</html>