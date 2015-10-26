<html>
<head>
<title>MyHomestay 标签调用说明</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/admin/Admin_STYLE.CSS" rel="stylesheet" type="text/css">
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0" class="bgcolor">

<br>
<table width="90%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center"> 
    <td height=25 colspan="2" class="topbg"><strong>系统模版标记说明</strong> 
  <tr> 
<td height=23 colspan="2" class="tdbg"><p><strong><font color=#0000ff>主模版（即系统首页，只对index.asp文件有效）：</font><br>
        注意</strong>：请使用正确的标签参数，否则可能出现不可预知的错误（一般为提示为下标越界）<br>
        　　　 红色部分为3.1版本新增或与3.0版本不同参数的标签。</p></td>
  </tr>
  <tr> 
    <td width="25%" height=23 class="tdbg"><p>$show_log(参数1,参数2,参数3,参数4,参数5,参数6,参数7,参数8,参数9)$<br>
      </p></td>
    <td width="75%" class="tdbg">此标记调用文章标题列表等信息。参数说明如下：<br>
      参数1：调用文章条数。<br>
      参数2：文章标题长度，以字符为单位，超过部分将不会显示。<br>
      参数3：排序方法，为1按发布时间，为2按点击数，为3按回复数。<br>
      参数4：是否精华，为1调用所有文章，为2调用精华文章。<br>
      参数5：调用多少天内的文章，以天为单位。<br>
      参数6：文章分类id，为0则调用所有分类的文章。<br>
      参数7：是否显示文章系统分类名，为1显示，为0不显示。<br>
      参数8：是否显示文章专题名，为1显示，为0不显示。<br>
      参数9：显示信息，1为显示发布时间和用户，2为显示发布时间，3为显示发布用户，4为显示发布用户和点击数，5为显示点击数，6为显示发布日期和用户，7为显示发布日期，8为显示回复数，0为不显示。</td>
  </tr>
  <tr> 
<td width="25%" height=23 class="tdbg"><p><font color="#FF0000">$show_userlog(参数1,参数2,参数3,参数4,参数5,参数6)$</font><br>
      </p></td>
    <td width="75%" class="tdbg"><font color="#FF0000">此标记调用某用户文章标题列表等信息。参数说明如下：<br>
      	参数1：userid。<br>
		参数2：调用文章条数。<br>
		参数3：文章标题长度，以字符为单位，超过部分将不会显示。<br>
		参数4：排序方法，为1按发布时间，为2按点击数，为3按回复数。<br>
		参数5：用户专题id，为0则调用该用户所有文章。<br>
		参数6：显示信息，1为显示发布日期，0为不显示。
	</font>
	</td>
  </tr>
  
  此标记调用文章标题列表等信息。参数说明如下：


  <tr> 
    <td height=23 class="tdbg">$show_placard$</td>
    <td height=23 class="tdbg">此标记显示系统公告。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_count$</td>
    <td height=23 class="tdbg"> 此标记站点统计信息。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_blogupdate(参数1)$</td>
    <td height=23 class="tdbg"> 此标记显示文章更新排行列表。参数1：调用条数。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_userlogin$</td>
    <td height=23 class="tdbg"> 此标记显示登录窗口。 </td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_comment(参数1,参数2)$</td>
    <td height=23 class="tdbg"> 此标记显示最新回复列表。参数1：调用条数；参数2：回复标题长度。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_subject(参数1)$</td>
    <td height=23 class="tdbg"> 此标记显示用户专题排行列表。参数1：调用条数。</td>
  </tr>
  <tr>
    <td height=23 class="tdbg">$show_bestblog(参数1)$</td>
    <td height=23 class="tdbg">此标记显示推荐用户。参数1：调用条数。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_newblogger(参数1)$</td>
    <td height=23 class="tdbg">此标记显示最新注册用户。参数1：调用条数。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_class(参数1)$</td>
    <td height=23 class="tdbg"> 此标记显示系统分类列表。参数1：横向显示时的每行条数，为0则竖向显示。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_bloger(参数1)$</td>
    <td height=23 class="tdbg"> 此标记显示用户列表。参数1：横向显示时的每行条数，为0则竖向显示。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_friends$</td>
    <td height=23 class="tdbg"> 此标记显示友情链接。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_search(参数1)$</td>
    <td height=23 class="tdbg"> 此标记显示搜索表单。参数1：为0横向显示，为1则竖向显示。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_cityblogger(参数1)$</td>
    <td height=23 class="tdbg">此标记显示同城用户搜索表单。参数1：为0横向显示，为1则竖向显示。<font color="#FF0000">此标签不能用于副模版。</font> 
    </td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_blogstar$</td>
    <td height=23 class="tdbg">此标记调用最新用户明星。</td>
  </tr>
    <tr> 
    <td height=23 class="tdbg">$show_blogstar2(参数1,参数2,参数3,参数4)$</td>
    <td height=23 class="tdbg">此标记调用最新用户明星。参数1：调用数目；参数2：每行显示数目；参数3：图片宽度；参数4：图片高度。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_newphoto(参数1,参数2,参数3,参数4)$</td>
    <td height=23 class="tdbg">此标签调用相册图片。参数1：调用条数；参数2：0为横向显示，1为纵向显示；参数3：图片宽度；参数4：图片高度。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_xml$</td>
    <td height=23 class="tdbg"> 此标记显示rss连接标志。</td>
  </tr>
  <tr> 
    <td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>副模版（对除index.asp外的其他系统页面有效，如HomestayNav.asp,listblogger.asp文件等）：</font></strong></td>
  </tr>
  <tr> 
    <td height=23 colspan="2" class="tdbg">包含所有主模版的标记，参数相同。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_list$</td>
    <td height=23 class="tdbg"> 重要，此标记显示其他系统次页面内容主体，不能去除，且只能在副摸版使用。</td>
  </tr>
</table>
<br>
<table width="90%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="border">
  <tr align="center"> 
    <td height=25 colspan="2" class="topbg"><strong>用户模版标记说明</strong> 
  <tr> 
    <td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>主模版：</font><br>
      </strong>主模版为页面的主体部分，包括css样式设置等，建议在Dreamweave或Frontpage中编辑，完成后将代码copy进来。 
    
    </td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_log$</td>
    <td height=23 class="tdbg"> 重要，此标记显示文章主体部分，包括评论等信息。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_placard$ </td>
    <td height=23 class="tdbg">此标记显示用户公告。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_calendar$</td>
    <td height=23 class="tdbg"> 此标记显示日历。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_newblog$</td>
    <td height=23 class="tdbg"> 此标记显示最新文章列表。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_comment$</td>
    <td height=23 class="tdbg"> 此标记显示最新回复列表。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_subject$</td>
    <td height=23 class="tdbg"> 此标记显示专题分类。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_subject_l$</td>
    <td height=23 class="tdbg"> 此标记横向显示专题分类。 </td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_newblog$</td>
    <td height=23 class="tdbg"> 此标记显示最新文章列表。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_newmessage$ </td>
    <td height=23 class="tdbg">此标记显示最新留言列表。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_info$</td>
    <td height=23 class="tdbg"> 此标记显示真实姓名称，统计信息等。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_login$</td>
    <td height=23 class="tdbg"> 此标记显示登录窗口。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_links$</td>
    <td height=23 class="tdbg"> 此标记显示链接信息。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_blogname$ </td>
    <td height=23 class="tdbg">此标记显示用户真实姓名称，若名称为空则显示用户id。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_search$</td>
    <td height=23 class="tdbg"> 此标记显示搜索表单。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_xml$</td>
    <td height=23 class="tdbg"> 此标记显示rss连接标志。</td>
  </tr>
  <tr> 
<td height=23 colspan="2" class="tdbg"><strong><font color=#0000ff>副模版：</font><br>
      </strong>副模版为显示文章内容部分。包括文章标题，发布时间，文章内容等信息的版面设置。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_topic$</td>
    <td height=23 class="tdbg"> 此标记显示表情图标，专题名，文章标题,。</td>
  </tr>
  <tr> 
    <td width="18%" height=23 class="tdbg">$show_loginfo$</td>
    <td width="82%" class="tdbg">此标记显示文章作者，发布时间等信息。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_logtext$</td>
    <td height=23 class="tdbg"> 此标记显示文章正文。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_more$</td>
    <td height=23 class="tdbg"> 此标记显示阅读全文(次数)，回复(次数)，引用链接。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_emot$</td>
    <td height=23 class="tdbg">此标记仅显示显示表情图标。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_author$</td>
    <td height=23 class="tdbg">此标记仅显示作者名。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_addtime$</td>
    <td height=23 class="tdbg">此标记仅显示发布时间。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_topictxt$</td>
    <td height=23 class="tdbg">此标记仅显示文章标题。</td>
  </tr>
  <tr> 
    <td height=23 class="tdbg">$show_blogzhai$</td>
    <td height=23 class="tdbg">此标记显示加入到文摘的连接。</td>
  </tr>
</table>
</body>
</html>