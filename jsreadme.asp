<html>
<head>
<style>
body {font-size:12px
}
table {font-size:12px
}
.style1 {color: #FF0000}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"></head>
<body>
<CENTER>
  <b>oBlog 3.0 首页调用测试与帮助文件</b>
</CENTER>
<p><BR>
  <BR>
  <BR>
  基本操作步骤:<BR>
  <font color=red> 1：在管理后台修改网站地址，必须为类似http://www.myhomestay.com.cn/blog/这样的完整目录；<BR>
  2：在需要调用的网页文件中加入调用代码，如：&lt;script src=ＸＸＸＸ&gt;&lt;/script></font></p>
<p>相关参数:<br>
  0: j=0 调用文章<BR>
  1: j=1 论坛统计&nbsp;<br>
  2: j=2 发帖文章用户排行&nbsp;<br>
  3: j=3 新注册用户排行<br>
  4: j=4 系统分类列表<br>
  5: j=5 个人专题排行列表<br>
  6: j=6 推荐用户排行列表<br>
  7: j=7 调用登陆窗口<br>
  8: j=8 调用系统公告<br>
  9: j=9 调用最新图片<br>
  10: j=10 调用最新的用户明星<br>
  n：显示多少个列表<br>
  <BR>
  <font color="#FF0000"><b>js首页调用演示</b></font><br>
  ============================================================================ 
</p>
<p>
<b>文章调用演示</b>
在页面插入的代码（以下为带系统分类名称的主题调用）:
<p> <font color="#FF0000">&lt;script src=&quot;js.asp?j=0&amp;classid=all&amp;subjectname=1&amp;classname=1&amp;tlen=16&amp;n=10&amp;sdate=90&amp;orders=2&amp;info=1&amp;action=1&amp;user=&quot;&gt;&lt;/script&gt;</font></p> 
<p> 相关参数<br>
  j=0:表示调用文章<br>
  classid :系统分类id，全部为all<br>
  subjectname : 0:为不调用 1:调用用户专题名称 <br>
  classname&nbsp;&nbsp; :0:为不调用 1:调用系统分类名称<br>
  tlen&nbsp;&nbsp;&nbsp; :标题长度<br>
  n　　 　:显示多少个标题<br>
  sdate 　:查询多少天内文章，1为当天<br>
  orders　: 排序方法，1为按照点击(最热文章)，2为按照时间(按最新回复时间),3:为按照评论数(最多回复文章)<br>
  info&nbsp;&nbsp;&nbsp; :1为显示发布时间和用户，2为显示发布时间，3为显示发布用户，4为显示发布用户和点击数，5为显示点击数，6为显示发布日期和用户，7为显示发布日期，8为显示回复数，0为不显示<br>
  action&nbsp; :1: 显示所有文章 2:显示精华文章<br>
  user &nbsp; :用户id号码（数字），显示该用户的文章(可以调用管理员文章，来做一个简单的新闻发布系统)<BR>
  <br>
  <script src="js.asp?j=0&classid=all&subjectname=1&classname=1&tlen=16&n=10&sdate=990&orders=2&info=1&action=1&user="></script>
  <BR>
  <BR>
  
  <strong>外教到京时间表 </strong><BR>
  在页面插入的代码:</p>
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=12&amp;classid=all&amp;subjectname=0&amp;classname=0&amp;tlen=1000&amp;n=10&amp;sdate=100&amp;orders=2&amp;info=0&amp;action=1&amp;user=&quot;&gt;&lt;/script&gt;</font></p>

<script src="js.asp?j=12&classid=all&subjectname=0&classname=0&tlen=1000&n=10&sdate=100&orders=2&info=0&action=1&user="></script>

<br><br>


  <strong>站点统计 </strong><BR>
  在页面插入的代码:</p>
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=1&quot;&gt;&lt;/script&gt;</font></p>

<script src="js.asp?j=1"></script>

<br><br>
<strong>发布文章前10名用户</strong><BR>
在页面插入的代码: 
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=2&amp;n=10&amp;order=0&quot;&gt;&lt;/script&gt;</font></p>

<script src="js.asp?j=2&n=10&order=0"></script>
<br>
<br><br>
<strong>总访问量排名前10名用户</strong><BR>
在页面插入的代码: 
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=2&amp;n=10&amp;order=1&quot;&gt;&lt;/script&gt;</font></p>

<script src="js.asp?j=2&n=10&order=1"></script>
<br>

<br>
<strong>新注册的10名会员 </strong><BR>
在页面插入的代码:
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=3&amp;n=10&quot;&gt;&lt;/script&gt;</font></p>

<script src="js.asp?j=3&n=10"></script>
<br>

<strong>新注册的30名中国家庭 </strong><BR>
在页面插入的代码:
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=11&amp;n=30&quot;&gt;&lt;/script&gt;</font></p>
<script src="js.asp?j=11&n=30"></script>
<br>


<p> <strong>系统文章分类</strong><BR>
在页面插入的代码:
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=4&quot;&gt;&lt;/script&gt;</font></p>

<script src="js.asp?j=4"></script>
<p><strong>用户分类</strong><br>
  在页面插入的代码: </p>
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=5&quot;&gt;&lt;/script&gt;</font></p>
<p> 
  <script src="js.asp?j=5"></script>
</p>
<p><strong>推荐用户排行前10名</strong><br>
  在页面插入的代码: </p>
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=6&amp;n=10&quot;&gt;&lt;/script&gt; </font></p>
 
<p>
  <script src="js.asp?j=6&n=10"></script>
</p>
<p><strong>显示登陆窗口</strong><br>
  在页面插入的代码: </p>
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=7&quot;&gt;&lt;/script&gt;<br>
  </font></p>
  <table width="200"><tr><td>
<script src="js.asp?j=7"></script></td></tr></table>
<p><strong>显示系统公告</strong><br>
在页面插入的代码: </p>
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=8&quot;&gt;&lt;/script&gt;</font></p>
<p> 
  <script src="js.asp?j=8"></script>
</p>
<p><strong>显示最新图片</strong><br>
  相关参数：<br>
  j=9　调用图片<br>
  n:调用条数<br>
  i:0为横向,1为纵向<br>
  w:宽度<br>
  h:高度 <br>
  在页面插入的代码:</p>
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=9&amp;n=4&amp;i=1&amp;w=100&amp;h=100&quot;&gt;&lt;/script&gt;</font></p>
<p> 
  <script src="js.asp?j=9&n=4&i=1&w=100&h=100"></script>
</p>
<p><strong>显示推荐用户</strong><br>
  相关参数：<br>
  j=10　调用推荐用户<br>
  n:调用条数<br>
  i:每行显示数目<br>
  w:宽度<br>
  h:高度 <br>
  在页面插入的代码:</p>
<p><font color="#FF0000">&lt;script src=&quot;js.asp?j=10&amp;n=2&amp;i=1&amp;w=130&amp;h=110&quot;&gt;&lt;/script&gt;</font></p>
<p> 
  <script src="js.asp?j=10&n=2&i=2&w=130&h=110"></script>
</p>
<p>&nbsp; </p>
</body>
</html> 
