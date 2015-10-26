<%

'全局常量或变量、参数定义程序段
'重要以下两个选项，每次修改文件都需认真查看

Const blogdir="/"
'blog程序所在目录,非常重要,默认为根目录,如为blog目录请改为/blog/

Const f_ext="html" 
'生成的文章静态文件后缀,可以为htm,html,shtml,asp四种格式

Const cookies_name="homestay_v20" 	'cookies名,一般无须修改
Const cookies_domain="" 	'cookeies域名根,一般留空
Const cache_name_user="Homestay" 	'系统缓存名前缀，一般无须修改
Const SYSFOLDER_ADMIN="admin"	'该目录名称将被作为系统禁止注册的用户名使用

const upload_dir="HomestayUPFiles"	'默认上传目录，为空则使用用户所在目录，若需要修改为其他目录，请手工建目录
const is_unamedir=1	'是否使用username作用户目录，１为使用，０为关闭
const en_nameisnum=1	'是否允许全数字的用户名,1为允许,0为不允许
const logfilepath=0	'文章文件路径，0为根目录，1为"/archives/年份/"目录  （=重要的设置wl=；）
const str_htmlfilt=""
'const str_htmlfilt="1|2|3|4|5|6|7|8|9|0|一|二|三|四|五|六|七|八|九|零|壹|贰|叁|肆|伍|陆|柒|捌|玖|拾|" '自定义过滤的html代码，必须以|结束，格式如aaa|bbb|，则过滤aaa和bbb标签(对评论和留言有效)

Const EN_PHOTO=1 	'相册开关,1为开启,0为关闭
Const EN_TAGS=1	 	'TAG开关
Const EN_COUNT = 1	'是否启用站点统计,访问量大的站点可以关闭次项目
Const is_debug=1 	'是否开启调试模式,1为开启，0为关闭
Const is_relativepath=0 '站内连接路径参数，1为相对路径，0为绝对路径

const P_TAGS_SPLIT=" "	 'TAG Split
const P_TAGS_MAX=10	 	 'TAGS IN ONE BLOG
const P_TAGS_DESC="自创关键字"	 'TAG DESC
const P_TAGS_ICON="TAG Keyword"	 'TAG DESC
const P_TAGS_CLOUD=1	 '1:CLOUD;2:LIST DESC
const P_TAGS_PerLine=8	 
const P_TAGS_SYSURL="tags.asp"
const P_TRACKBACK_MAX=5

Const P_Site_Hours=0	'服务器与时差设置
Const P_BLOG_UPDATEPAUSE=2 	'每生成5篇文章的暂停时间(1~100)
const true_domain=0	'是否启用真实二级域名,需OblogDns组件支持

dim blogurl,str_domain
if true_domain=1 then
	'真实域名必填，设置blog程序绝对路径
	blogurl="http://www.myhomestay.com.cn/"
	str_domain=",custom_domain"
else
	blogurl=blogdir
end if
%>