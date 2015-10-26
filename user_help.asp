<!--#include file="user_top.asp"-->
<!--#include file="inc/class_blog.asp"-->
<div id="main">
  <div class="submenu">
  	<div class="side_c1 side11"></div>
    <div class="side_c2 side21"></div>
	<div class="submenu_content">
	</div>
  </div>
  <div class="content">
  	<div class="content_top">
		  	<div class="side_d1 side11"></div>
		    <div class="side_d2 side21"></div>
			帮助
	</div>
	
    <div class="content_body">
	<%call main%>
	</div>
	
    <div class="content_bot">
		  	<div class="side_e1 side12"></div>
   			<div class="side_e2 side22"></div>
 	</div>
	
  </div>
</div>   
  
  <div id="bottom"><%=oblog.user_copyright%></div>

</body>
</html>
<%sub main%>
<h1>用户后台帮助</h1>
<div id="list">
	<h1><a name="1"></a>我的blog出现版面错乱，挪位怎么办？ </h1> 
	<ul class="list_edit">
	1、检查模版是否正常。<br />
	2、检查文章是否有不标准的htm代码。<br />
	3、对文章部分显示字数和排版进行微调，达到正常为止。<br />
	4、建议使用部分显示标签进行首页部分显示文章的排版。
	</ul>
	<h1><a name="2" id="2"></a>模版不小心改坏了怎么办？ </h1> 
	<ul class="list_edit">
	重新选择一个默认模版即可。注：会将原来改过的模版覆盖掉，建议先备份模版。<br />
	</ul>
	<h1><a name="3" id="3"></a>文章发布了，但为什么首页没有显示？ </h1> 
	<ul class="list_edit">
	请用更新按钮，重新发布站点首页。<br />
	</ul>
	<h1><a name="4" id="4"></a>为什么无法上传文件？ </h1> 
	<ul class="list_edit">
	1、是您上传的文件大小超过了系统设定值，请压缩图片或文件。<br />
	2、您的上传空间已满，请整理您的上传文件。<br />
	3、您没有上传权限，请联系管理员。 <br />
	</ul>
	<h1><a name="5" id="5"></a>一篇文章最多能写多少字？ </h1> 
	<ul class="list_edit">
	因受到数据库字段长度的限制，一篇文章请不要超过6万个英文字符，即：3万个中文字符。<br />
	</ul>
	<h1><a name="6" id="6"></a>可视化编辑器支持哪几种浏览器？ </h1> 
	<ul class="list_edit">
	oBlog集成的编辑器可以支持5.0以上版本的ie全系列浏览器，在mozilla,firefox浏览器有部分功能无法使用，在opear浏览器环境无法使用。<br />
	建议非ie浏览器使用ubb编辑器发布文章，在站点配置菜单可以选择后台默认编辑器。
	</ul>
	<h1><a name="7" id="7"></a>为什么我无法登陆管理后台？ </h1> 
	<ul class="list_edit">
	1、请确认用户名和密码输入正确。<br />
	2、登陆系统需要cookies环境，请检查浏览器的cookies是否关闭。<br />
	3、请联系系统管理员。<br />
	</ul>
	<h1><a name="9" id="9"></a>如何修改我的个性模版？ </h1> 
	<ul class="list_edit">
	请先选择一个喜欢的默认模版，然后选择<strong>摸板设置</strong>菜单进行操作。<br />
	建议将代码拷贝到Dreamweaver 或者Frontpage编辑 
	<p>注意：本系统分为主模版和副模版，主模版为网站的整体结构，副模版的修改只对文章主体部分起作用，也就是对标签$show_log$起作用，具体调用标签如下，您可以灵活运用，做个个性的模版（建议修改前先备份模版）</p>
	<p>用户模版标记说明 <br />
	  主模版： <br />
	  $show_log$ 重要，此标记显示文章主体部分，包括评论等信息。 <br />
	  $show_placard$ 此标记显示用户公告。 <br />
	  $show_calendar$ 此标记显示日历。 <br />
	  $show_newblog$ 此标记显示最新文章列表。 <br />
	  $show_comment$ 此标记显示最新回复列表。 <br />
	  $show_subject$ 此标记显示专题分类。 <br />
	  $show_subject_l$ 此标记横向显示专题分类。 <br />
	  $show_newblog$ 此标记显示最新文章列表。 <br />
	  $show_newmessage$ 此标记显示最新留言列表。 <br />
	  $show_info$ 此标记显示真实姓名称，统计信息等。 <br />
	  $show_login$ 此标记显示登陆窗口 <br />
	  $show_links$ 此标记显示链接信息 <br />
	  $show_blogname$ 此标记显示用户真实姓名称，若名称为空则显示用户id。 <br />
	  $show_search$ 此标记显示搜索表单 <br />
	  $show_xml$ 此标记显示rss连接标志。 <br />
	  副模版： <br />
	  $show_topic$ 此标记显示文章题目。 <br />
	  $show_loginfo$ 此标记显示文章作者，发布时间等信息。 <br />
	  $show_logtext$ 此标记显示文章正文。 <br />
	  $show_more$ 此标记显示阅读全文，引用等链接。 <br />
	  $show_emot$ 此标记仅显示显示表情图标。<br />
	  $show_author$ 此标记仅显示作者名。<br />
	  $show_addtime$ 此标记仅显示发布时间。<br />
	  $show_topictxt$ 此标记仅显示文章标题。</p>
	注意：若不小心将模版改坏，可以重新选择默认模版进行恢复。<br />
	</ul>
	<h1><a name="tb">什么是trackback ping(引用通告)？</a></h1>
	<ul class="list_edit">
	一、“引用通告”是什么？<br />    　　“引用通告”简单的说，就是如果你写的文章是根据其他人Blog中的文章而做出的延伸或评论，你可以通知对方你针对他的文章写了东西，这就需要用到引用通告。<br />
　　在以往我们的经验当中，您对他人文章文章的评论只能在他人文章后通过回复进行，这样做让我们只能在别人的地盘上活动，而不能自己掌握自己的发布的言论，这就带来一些麻烦。<br />
　　第一，您发布在他人文章后面的评论您没有办法再进行修改。如果张三的Blog上有一篇我感兴趣的文章，您在这篇文章下发布自己的评论，但您的评论只能存在于张三的Blog上，您无法再修改增删这篇评论。<br />
　　第二，您希望获得张三关注的文章只能采取在自己的Blog中写了一篇和张三类似的文章，您希望张三能来看一看我写的这篇文章，这时您就必须到张三的Blog的那篇文章下发一篇回复，同时把您想让他看的那篇文章地址贴上去。
有了引用通告，您就可以完全不需要这样麻烦了，完全可以在自己的BLOG里进行操作，彻底享受自己掌握自己言论的主动权。<br />
　　通过引用通告，您就可以在自己的Blog中发布文章，同时把自己这篇文章的地址发到张三的Blog的那篇文章上去。
在自己的地盘上引用张三的文章，然后通知他，“嘿，您的文章被我评论了”，这就是“引用通告”。<br />

　　同样的，当别人引用您的文章的时候，系统也可以接收对方的请求并进行记录，这样您可以查看来源地址，看看对方是如何评论您的文章的。<br />
<br />
二、如何使用“引用通告”<br />
　　1、找到你要评论的Blog文章的“引用通告”地址，一般在文章下方有“引用通告”项，点击进入后可以看到地址，或者有的Blog直接在文章下方显示，把地址复制下来；<br />
　　2、进入自己的Blog发布新文章，在下方有“引用通告”栏目，将要评论的Blog的“引用通告”地址粘贴在这里。每行您可以粘帖一个。<br />
　　3、发布自己的文章后，系统会自动向目标地址发送引用申请，之后你会在要评论的Blog中看到你的Blog文章地址。<br />
	<br />
　　引用通告并不神秘和复杂，您可以先和朋友互相试验一下，相信您很快就会发现引用通告给您带来的全新感受。
	</ul>
	<h1><a name="tag">什么是Tag？</a></h1>
	<ul class="list_edit">
一、什么是标签（TAG）？<br />
　　简单的说,标签就是一篇文章的"关键词"。您可以将文章文章或者照片，选择一个或多个词语（标签）来标记，这样一来，凡是我们用户网站上使用该词语的文章自动成为一个列表显示。<br />

二、使用标签的好处：<br />
　　1、您添加标签的文章就会被直链接到网站相应标签的页面，这样浏览者在访问相关标签时，就有可能访问到您的文章，增加了您的文章被访问的机会。<br />
　　2、您可以很方便地查找到与您使用了同样标签的文章，延伸您文章的视野；可以方便地查找到与您使用了同样标签的作者，作为志同道合的朋友，您可以将他们加为好友或友情用户，扩大您的朋友圈。 <br />
　　3、增加标签的方式完全由您自主决定，不受任何的限制，不用受网站系统分类和自己原有文章分类的限制，便于信息的整理、记忆和查找。<br />

三、如何使用标签?<br />
　　例如：您写了一篇到北京旅游的文章，按照文章提到的内容，您可以给这篇文章加上：<br />
　　北京旅游,天安门,长城,故宫<br />
　　等几个标签，当浏览者想搜索关于长城的文章时，浏览者会点击标签：长城，从而看到所有关于长城的文章，方便了浏览者查找文章，同时您也可用此方法找到和您同样喜欢的人，以便一起相互交流等等。<br />

四．如何添加“好”的Tag？ <br />
　　1． Tag应该要能够体现出自己的特色，并且是大家经常采用的熟悉的词语。<br /> 
　　2．用词尽量简单精炼,词语字数不要太长，两三个字的词语就可以了，尽量是有意义的词汇，不要使用一些只作为装饰的符号，如｛｝等。<br />
　　3．不要使用一些语义比较弱的词汇，如“我的家”，“图片”等。<br />
	</ul>
	<h1><a name="edit">编辑器帮助</a></h1>
	<ul class="list_edit">
	 oBlog集成了可视化编辑器，可以非常方便地编辑您的文章。<br />
      您可以把鼠标悬停在操作按纽上得到快速提示，也可以对照下面的表格得到详细说明。<br />
	<hr />
	常用： <br />
	<img src="images/img.gif" > 在当前位置插入或者上传图片。<br />
	<img src="images/smiley.gif" > 在当前位置插入表情图标。<br />
	<img src="images/file.gif" > 上传文件。 <br />
	<img src="images/swf.gif" > 在当前位置插入flash动画。<br />
	<img src="images/wmv.gif" > 在当前位置插入media play媒体。<br />
	<img src="images/rm.gif" > 在当前位置插入rm媒体。 <br />
	<img src="images/wlink.gif" > 为选定的内容增加或修改超级连接。<br />
		  <img src="images/unlink.gif" > 取消选定内容的超级连接。 
		  <hr size="1" noshade>
	格式： <br />
	<img src="images/fgcolor.gif" > 设定字体颜色。<br />
	<img src="images/fbcolor.gif" > 设定字体背景颜色。<br />
	<img src="images/bold.gif" > 设定选定文字的格式为粗体。<br />
	<img src="images/italic.gif" > 设定选定文字的格式为斜体。<br />
	<img src="images/underline.gif" > 设定选定文字的格式为下划线。<br />
	<img src="images/superscript.gif" > 设定选定文字的格式为上标。<br />
	<img src="images/subscript.gif" > 设定选定文字的格式为下标。<br />
	<img src="images/strikethrough.gif" > 设定选定文字的格式为删除线。<br />
	<img src="images/removeformat.gif" > 取消选定文字的所有格式。<br />
	<img src="images/aleft.gif" > 设定选定内容为左对齐。<br />
	<img src="images/center.gif" > 设定选定内容为中对齐。<br />
	<img src="images/aright.gif" > 设定选定内容为右对齐。<br />
	<img src="images/numlist.gif" > 设定选定文字为编号格式。<br />
	<img src="images/bullist.gif" > 设定选定文字为项目符号格式。<br />
	<img src="images/outdent.gif" > 减少选定内容的缩进量。<br />
	<img src="images/indent.gif" > 增加选定内容的缩进量。
	<hr size="1" noshade>
	其他： <br />
	<img src="images/specialchar.gif" > 插入特殊字符。<br />
	<img src="images/replace.gif" > 查找替换文字。<br />
	<img src="images/cleancode.gif"> 清理垃圾htm代码，从word粘贴过来的内容建议使用。<br />
	<img src="images/selectAll.gif" > 选择文章全部内容。<br />
	<img src="images/cut.gif" > 剪切选定的文章内容。<br />
	<img src="images/copy.gif" > 复制选定的文章内容。<br />
	<img src="images/paste.gif" > 从剪贴板粘贴内容。<br />
	<img src="images/undo.gif" > 撤消上一步操作。<br />
	<img src="images/redo.gif" > 恢复撤消的操作。 <br />
	<img src="images/hr.gif" > 在当前位置插入水平线。<br />
	<img src="images/table.gif" > 插入表格。<br />
	<img src="images/insertrow.gif" > 在当前位置的表格插入行。<br />
	<img src="images/deleterow.gif" > 删除当前表格位置的行。<br />
	<img src="images/insertcolumn.gif" > 在当前位置的表格插入列。<br />
	<img src="images/deletecolumn.gif" > 删除当前表格位置的列。 <br />
	<img src="images/code.gif" > 设定选定内容为代码样式。<br />
	<img src="images/quote.gif" > 设定选定内容为引用样式。<br />
	</ul>
</div>
<%end sub%>