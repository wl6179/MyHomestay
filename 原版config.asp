<%

'ȫ�ֳ����������������������
'��Ҫ��������ѡ�ÿ���޸��ļ���������鿴

Const blogdir="/"
'blog��������Ŀ¼,�ǳ���Ҫ,Ĭ��Ϊ��Ŀ¼,��ΪblogĿ¼���Ϊ/blog/

Const f_ext="html" 
'���ɵ����¾�̬�ļ���׺,����Ϊhtm,html,shtml,asp���ָ�ʽ

Const cookies_name="oblog312" 	'cookies��,һ�������޸�
Const cookies_domain="" 	'cookeies������,һ������
Const cache_name_user="oblog" 	'ϵͳ������ǰ׺��һ�������޸�
Const SYSFOLDER_ADMIN="ADMIN"	'��Ŀ¼���ƽ�����Ϊϵͳ��ֹע����û���ʹ��

const upload_dir="UploadFiles"	'Ĭ���ϴ�Ŀ¼��Ϊ����ʹ���û�����Ŀ¼������Ҫ�޸�Ϊ����Ŀ¼�����ֹ���Ŀ¼
const is_unamedir=1	'�Ƿ�ʹ��username���û�Ŀ¼����Ϊʹ�ã���Ϊ�ر�
const en_nameisnum=0	'�Ƿ�����ȫ���ֵ��û���,1Ϊ����,0Ϊ������
const logfilepath=1	'�����ļ�·����0Ϊ��Ŀ¼��1Ϊ"/archives/���/"Ŀ¼
const str_htmlfilt="" '�Զ�����˵�html���룬������|��������ʽ��aaa|bbb|�������aaa��bbb��ǩ(�����ۺ�������Ч)

Const EN_PHOTO=1 	'��Ὺ��,1Ϊ����,0Ϊ�ر�
Const EN_TAGS=1	 	'TAG����
Const EN_COUNT = 1	'�Ƿ�����վ��ͳ��,���������վ����Թرմ���Ŀ
Const is_debug=1 	'�Ƿ�������ģʽ,1Ϊ������0Ϊ�ر�
Const is_relativepath=0 'վ������·��������1Ϊ���·����0Ϊ����·��

const P_TAGS_SPLIT=" "	 'TAG Split
const P_TAGS_MAX=10	 	 'TAGS IN ONE BLOG
const P_TAGS_DESC="���±�ǩ"	 'TAG DESC
const P_TAGS_ICON="TAG"	 'TAG DESC
const P_TAGS_CLOUD=1	 '1:CLOUD;2:LIST DESC
const P_TAGS_PerLine=8	 
const P_TAGS_SYSURL="tags.asp"
const P_TRACKBACK_MAX=5

Const P_Site_Hours=0	'��������ʱ������
Const P_BLOG_UPDATEPAUSE=2 	'ÿ����5ƪ���µ���ͣʱ��(1~100)
const true_domain=0	'�Ƿ�������ʵ��������,��OblogDns���֧��

dim blogurl,str_domain
if true_domain=1 then
	'��ʵ�����������blog�������·��
	blogurl="http://www.aaa.com/blog/"
	str_domain=",custom_domain"
else
	blogurl=blogdir
end if
%>