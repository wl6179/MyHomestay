if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ob_calendar]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[ob_calendar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_admin]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_admin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_blogstar]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_blogstar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_blogteam]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_blogteam]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_calendar]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_calendar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_comment]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_comment]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_friend]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_friend]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_log]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_log]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_logclass]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_logclass]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_message]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_message]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_notdownload]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_notdownload]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_pm]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_pm]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_setup]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_setup]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_subject]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_subject]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_sysskin]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_sysskin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_tags]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_tags]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_trackback]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_trackback]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_upfile]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_upfile]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_user]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_user]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_userclass]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_userclass]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_userdir]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_userdir]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_userskin]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_userskin]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[oblog_usertags]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[oblog_usertags]
GO

CREATE TABLE [dbo].[oblog_admin] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[password] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[lastloginip] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[lastlogintime] [smalldatetime] NULL ,
	[logintimes] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_blogstar] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[userid] [int] NULL ,
	[userurl] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[picurl] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[info] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[addtime] [smalldatetime] NULL ,
	[blogname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[ispass] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_blogteam] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[mainuserid] [int] NOT NULL ,
	[otheruserid] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_calendar] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[userid] [int] NULL ,
	[calendar_month] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_comment] (
	[commentid] [int] IDENTITY (1, 1) NOT NULL ,
	[mainid] [int] NOT NULL ,
	[comment] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[addtime] [smalldatetime] NULL ,
	[comment_user] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[commenttopic] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[userid] [int] NOT NULL ,
	[addip] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[homepage] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[isguest] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_friend] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[userid] [int] NOT NULL ,
	[friendid] [int] NOT NULL ,
	[addtime] [smalldatetime] NOT NULL ,
	[isblack] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_log] (
	[logid] [int] IDENTITY (1, 1) NOT NULL ,
	[topic] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[logtext] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[userid] [int] NULL ,
	[authorid] [int] NULL ,
	[author] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[addtime] [smalldatetime] NULL ,
	[truetime] [smalldatetime] NULL ,
	[iis] [int] NULL ,
	[face] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[commentnum] [int] NULL ,
	[classid] [int] NULL ,
	[subjectid] [int] NULL ,
	[showword] [int] NULL ,
	[ishide] [int] NULL ,
	[ispassword] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[trackbacknum] [int] NULL ,
	[passcheck] [int] NULL ,
	[istop] [int] NULL ,
	[isbest] [int] NULL ,
	[isencomment] [int] NULL ,
	[blog_password] [int] NULL ,
	[tburl] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[logfile] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[isdraft] [int] NULL ,
	[issave] [int] NULL ,
	[subjectfile] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[logTags] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[logTagsId] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[logType] [int] NULL ,
	[Abstract] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[Abstract_Auto] [int] NULL ,
	[EditorType] [int] NULL ,
	[filename] [varchar] (30) COLLATE Chinese_PRC_CI_AS NULL ,
	[logpics] [text] COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_logclass] (
	[classid] [int] IDENTITY (1, 1) NOT NULL ,
	[id] [int] NULL ,
	[classname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[classlognum] [int] NULL ,
	[ordernum] [int] NULL ,
	[orderid] [int] NULL ,
	[parentid] [int] NULL ,
	[parentpath] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[depth] [int] NULL ,
	[rootid] [int] NULL ,
	[child] [int] NULL ,
	[readme] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[previd] [int] NULL ,
	[nextid] [int] NULL ,
	[idType] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_message] (
	[messageid] [int] IDENTITY (1, 1) NOT NULL ,
	[userid] [int] NULL ,
	[messagetopic] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[message] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[addtime] [smalldatetime] NULL ,
	[message_user] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[addip] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[homepage] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[isguest] [int] NULL ,
	[issave] [int] NULL ,
	[messagefile] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_notdownload] (
	[notdown] [image] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_pm] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[sender] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[incept] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[topic] [nvarchar] (100) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[content] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[isreaded] [int] NOT NULL ,
	[addtime] [smalldatetime] NOT NULL ,
	[delR] [int] NOT NULL ,
	[delS] [int] NOT NULL ,
	[isguest] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_setup] (
	[ID] [int] NOT NULL ,
	[site_name] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[site_title] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[site_path] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[site_domain] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[site_keyword] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[master_email] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[site_copyright] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[show_log_num] [int] NULL ,
	[is_log_check] [int] NULL ,
	[is_comment_check] [int] NULL ,
	[is_enguest_comment] [int] NULL ,
	[is_secondary_domain] [int] NULL ,
	[site_timediff] [int] NULL ,
	[get_log_num] [int] NULL ,
	[is_enreg] [int] NULL ,
	[is_admin_chkreg] [int] NULL ,
	[time_index_update] [int] NULL ,
	[site_placard] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[site_friends] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[user_placard] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[log_count] [int] NULL ,
	[comment_count] [int] NULL ,
	[message_count] [int] NULL ,
	[user_count] [int] NULL ,
	[is_nickname_rep] [int] NULL ,
	[oneday_update] [int] NULL ,
	[show_imgw_num] [int] NULL ,
	[show_img_mouse] [int] NULL ,
	[is_log_profilt] [int] NULL ,
	[default_userdir] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[is_vip_prosee] [int] NULL ,
	[max_badstr_num] [int] NULL ,
	[upfile_user_en] [int] NULL ,
	[upfile_user_type] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[upfile_user_onesize] [int] NULL ,
	[upfile_user_maxsize] [int] NULL ,
	[upfile_vip_en] [int] NULL ,
	[upfile_vip_type] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[upfile_vip_onesize] [int] NULL ,
	[upfile_vip_maxsize] [int] NULL ,
	[upfile_admin_en] [int] NULL ,
	[upfile_admin_type] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[upfile_admin_onesize] [int] NULL ,
	[upfile_admin_maxsize] [int] NULL ,
	[show_listuser_num] [int] NULL ,
	[show_comment_asc] [int] NULL ,
	[reg_text] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[is_need_classid] [int] NULL ,
	[add_pitchtime] [int] NULL ,
	[log_badstr] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[reg_badstr] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[lockip] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[pm_num] [int] NULL ,
	[enchangdomain] [int] NULL ,
	[dis_refutime] [int] NULL ,
	[ad_usermanage] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[is_code_reg] [int] NULL ,
	[is_code_login] [int] NULL ,
	[is_code_comment] [int] NULL ,
	[is_code_addblog] [int] NULL ,
	[is_code_pm] [int] NULL ,
	[upset_uptype] [int] NULL ,
	[upset_isDraw] [int] NULL ,
	[upset_Drawtext] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[upset_Drawfontsize] [int] NULL ,
	[upset_Drawcolor] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[upset_Drawfont] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[upset_DrawFontBold] [int] NULL ,
	[upset_Drawpic] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[upset_DrawGraph] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[upset_Drawpiccolor] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[upset_DrawHight] [int] NULL ,
	[upset_DrawWidth] [int] NULL ,
	[upset_DrawXYType] [int] NULL ,
	[max_log_word] [int] NULL ,
	[max_comment_word] [int] NULL ,
	[is_searchlogtext] [int] NULL ,
	[get_photonum] [int] NULL ,
	[show_photonum] [int] NULL ,
	[tag_badstr] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[index_to_file] [int] NULL ,
	[loadingstr] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[df_filename] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_subject] (
	[subjectid] [int] IDENTITY (1, 1) NOT NULL ,
	[userid] [int] NULL ,
	[subjectname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[subjectlognum] [int] NULL ,
	[ordernum] [int] NULL ,
	[subjectType] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_sysskin] (
	[id] [int] IDENTITY (1, 1) NOT NULL ,
	[sysskinname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[skinMain] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[skinShowlog] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[skinauthor] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[isdefault] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_tags] (
	[TagId] [int] IDENTITY (1, 1) NOT NULL ,
	[Name] [varchar] (255) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[iClass] [int] NOT NULL ,
	[iNum] [int] NOT NULL ,
	[iType] [int] NOT NULL ,
	[iState] [int] NOT NULL ,
	[LastUpdate] [smalldatetime] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_trackback] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[logid] [int] NULL ,
	[topic] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[tbuser] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[url] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[addtime] [smalldatetime] NULL ,
	[Excerpt] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[VisitorUrl] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_upfile] (
	[fileid] [int] IDENTITY (1, 1) NOT NULL ,
	[userid] [int] NOT NULL ,
	[file_name] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[file_path] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[file_ext] [nvarchar] (10) COLLATE Chinese_PRC_CI_AS NULL ,
	[file_size] [int] NULL ,
	[file_readme] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[isphoto] [smallint] NOT NULL ,
	[addtime] [smalldatetime] NOT NULL ,
	[file_remark] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[issave] [smallint] NOT NULL ,
	[photofile] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[logid] [int] NOT NULL ,
	[isPower] [int] NOT NULL ,
	[photoExif] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[sysClassId] [int] NOT NULL ,
	[UserClassId] [int] NOT NULL ,
	[photo_width] [int] NOT NULL ,
	[photo_height] [int] NOT NULL ,
	[viewNum] [int] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_user] (
	[userid] [int] IDENTITY (1, 1) NOT NULL ,
	[username] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[user_domain] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[user_domainroot] [nvarchar] (20) COLLATE Chinese_PRC_CI_AS NULL ,
	[user_dir] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[useremail] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[password] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NOT NULL ,
	[nickname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[truename] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[sex] [int] NULL ,
	[adddate] [smalldatetime] NULL ,
	[province] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[city] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[address] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[qq] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[tel] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[lastloginip] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[lastlogintime] [smalldatetime] NULL ,
	[logintimes] [int] NULL ,
	[lockuser] [int] NULL ,
	[blogname] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[birthday] [smalldatetime] NULL ,
	[question] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[answer] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[msn] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[homepage] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[job] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[user_isbest] [int] NULL ,
	[siteinfo] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[user_skin_main] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[user_skin_showlog] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[log_count] [int] NULL ,
	[user_level] [int] NULL ,
	[user_placard] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[user_links] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[user_showlog_num] [int] NULL ,
	[hideurl] [int] NULL ,
	[user_shownewlog_num] [int] NULL ,
	[user_shownewcomment_num] [int] NULL ,
	[user_shownewmessage_num] [int] NULL ,
	[user_siterefu_num] [int] NULL ,
	[comment_count] [int] NULL ,
	[message_count] [int] NULL ,
	[blog_password] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[user_upfiles_max] [int] NULL ,
	[user_classid] [int] NULL ,
	[user_upfiles_size] [int] NULL ,
	[user_upfiles_num] [int] NULL ,
	[user_showlogword_num] [int] NULL ,
	[en_blogteam] [int] NULL ,
	[comment_isasc] [int] NULL ,
	[defaultskin] [int] NULL ,
	[bak_skin1] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[bak_skin2] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[isubbedit] [int] NULL ,
	[vipdate] [smalldatetime] NULL ,
	[user_folder] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[user_icon1] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[user_icon2] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[user_info] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[user_css] [text] COLLATE Chinese_PRC_CI_AS NULL ,
	[user_cookie] [varchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[user_photorow_num] [int] NULL ,
	[custom_domain] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_userclass] (
	[classid] [int] IDENTITY (1, 1) NOT NULL ,
	[id] [int] NULL ,
	[classname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[classlognum] [int] NULL ,
	[ordernum] [int] NULL ,
	[orderid] [int] NULL ,
	[parentid] [int] NULL ,
	[parentpath] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[depth] [int] NULL ,
	[rootid] [int] NULL ,
	[child] [int] NULL ,
	[readme] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[previd] [int] NULL ,
	[nextid] [int] NULL ,
	[idType] [int] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_userdir] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[userdir] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[is_default] [int] NULL ,
	[dirdomain] [varchar] (255) COLLATE Chinese_PRC_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_userskin] (
	[ID] [int] IDENTITY (1, 1) NOT NULL ,
	[userskinname] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[skinmain] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[skinshowlog] [ntext] COLLATE Chinese_PRC_CI_AS NULL ,
	[skinauthor] [nvarchar] (50) COLLATE Chinese_PRC_CI_AS NULL ,
	[skinauthorurl] [nvarchar] (255) COLLATE Chinese_PRC_CI_AS NULL ,
	[isdefault] [int] NULL ,
	[skinpic] [nvarchar] (250) COLLATE Chinese_PRC_CI_AS NULL ,
	[ispass] [int] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[oblog_usertags] (
	[UserTagId] [int] IDENTITY (1, 1) NOT NULL ,
	[TagId] [int] NOT NULL ,
	[UserId] [int] NOT NULL ,
	[logId] [int] NOT NULL ,
	[iNum] [int] NOT NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[oblog_admin] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_admin] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_blogstar] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_blogstar] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_blogteam] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_blogteam] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_calendar] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_calendar] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_comment] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_comment] PRIMARY KEY  CLUSTERED 
	(
		[commentid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_friend] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_friend] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_log] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_log] PRIMARY KEY  CLUSTERED 
	(
		[logid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_logclass] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_logclass] PRIMARY KEY  CLUSTERED 
	(
		[classid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_message] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_message] PRIMARY KEY  CLUSTERED 
	(
		[messageid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_pm] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_pm] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_setup] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_setup] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_subject] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_subject] PRIMARY KEY  CLUSTERED 
	(
		[subjectid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_sysskin] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_sysskin] PRIMARY KEY  CLUSTERED 
	(
		[id]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_tags] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_tags] PRIMARY KEY  CLUSTERED 
	(
		[TagId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_trackback] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_trackback] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_upfile] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_upfile] PRIMARY KEY  CLUSTERED 
	(
		[fileid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_user] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_user] PRIMARY KEY  CLUSTERED 
	(
		[userid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_userclass] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_userclass] PRIMARY KEY  CLUSTERED 
	(
		[classid]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_userdir] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_userdir] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_userskin] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_userskin] PRIMARY KEY  CLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_usertags] WITH NOCHECK ADD 
	CONSTRAINT [PK_oblog_usertags] PRIMARY KEY  CLUSTERED 
	(
		[UserTagId]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[oblog_blogstar] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_blogstar_addtime] DEFAULT (getdate()) FOR [addtime],
	CONSTRAINT [DF_oblog_blogstar_ispass] DEFAULT (0) FOR [ispass]
GO

ALTER TABLE [dbo].[oblog_comment] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_comment_mainid] DEFAULT (0) FOR [mainid],
	CONSTRAINT [DF_oblog_comment_addtime] DEFAULT (getdate()) FOR [addtime],
	CONSTRAINT [DF_oblog_comment_userid] DEFAULT (0) FOR [userid],
	CONSTRAINT [DF_oblog_comment_isguest] DEFAULT (0) FOR [isguest]
GO

ALTER TABLE [dbo].[oblog_friend] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_friend_addtime] DEFAULT (getdate()) FOR [addtime],
	CONSTRAINT [DF_oblog_friend_isblack] DEFAULT (0) FOR [isblack]
GO

ALTER TABLE [dbo].[oblog_log] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_log_userid] DEFAULT (0) FOR [userid],
	CONSTRAINT [DF_oblog_log_authorid] DEFAULT (0) FOR [authorid],
	CONSTRAINT [DF_oblog_log_truetime] DEFAULT (getdate()) FOR [truetime],
	CONSTRAINT [DF_oblog_log_iis] DEFAULT (0) FOR [iis],
	CONSTRAINT [DF_oblog_log_commentnum] DEFAULT (0) FOR [commentnum],
	CONSTRAINT [DF_oblog_log_classid] DEFAULT (0) FOR [classid],
	CONSTRAINT [DF_oblog_log_subjectid] DEFAULT (0) FOR [subjectid],
	CONSTRAINT [DF_oblog_log_showword] DEFAULT (0) FOR [showword],
	CONSTRAINT [DF_oblog_log_ishide] DEFAULT (0) FOR [ishide],
	CONSTRAINT [DF_oblog_log_trackbacknum] DEFAULT (0) FOR [trackbacknum],
	CONSTRAINT [DF_oblog_log_passcheck] DEFAULT (1) FOR [passcheck],
	CONSTRAINT [DF_oblog_log_istop] DEFAULT (0) FOR [istop],
	CONSTRAINT [DF_oblog_log_isbest] DEFAULT (0) FOR [isbest],
	CONSTRAINT [DF_oblog_log_isencomment] DEFAULT (1) FOR [isencomment],
	CONSTRAINT [DF_oblog_log_blog_password] DEFAULT (0) FOR [blog_password],
	CONSTRAINT [DF_oblog_log_isdraft] DEFAULT (0) FOR [isdraft],
	CONSTRAINT [DF_oblog_log_issave] DEFAULT (0) FOR [issave],
	CONSTRAINT [DF__oblog_log__logTy__245D67DE] DEFAULT (0) FOR [logType],
	CONSTRAINT [DF__oblog_log__Abstr__25518C17] DEFAULT (1) FOR [Abstract_Auto],
	CONSTRAINT [DF__oblog_log__Edito__2645B050] DEFAULT (0) FOR [EditorType]
GO

ALTER TABLE [dbo].[oblog_logclass] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_logclass_id] DEFAULT (0) FOR [id],
	CONSTRAINT [DF_oblog_logclass_classlognum] DEFAULT (0) FOR [classlognum],
	CONSTRAINT [DF_oblog_logclass_ordernum] DEFAULT (0) FOR [ordernum],
	CONSTRAINT [DF_oblog_logclass_orderid] DEFAULT (0) FOR [orderid],
	CONSTRAINT [DF_oblog_logclass_parentid] DEFAULT (0) FOR [parentid],
	CONSTRAINT [DF_oblog_logclass_depth] DEFAULT (0) FOR [depth],
	CONSTRAINT [DF_oblog_logclass_rootid] DEFAULT (0) FOR [rootid],
	CONSTRAINT [DF_oblog_logclass_child] DEFAULT (0) FOR [child],
	CONSTRAINT [DF_oblog_logclass_previd] DEFAULT (0) FOR [previd],
	CONSTRAINT [DF_oblog_logclass_nextid] DEFAULT (0) FOR [nextid],
	CONSTRAINT [DF__oblog_log__idTyp__29221CFB] DEFAULT (0) FOR [idType]
GO

ALTER TABLE [dbo].[oblog_message] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_message_userid] DEFAULT (0) FOR [userid],
	CONSTRAINT [DF_oblog_message_addtime] DEFAULT (getdate()) FOR [addtime],
	CONSTRAINT [DF_oblog_message_isguest] DEFAULT (0) FOR [isguest],
	CONSTRAINT [DF_oblog_message_issave] DEFAULT (0) FOR [issave]
GO

ALTER TABLE [dbo].[oblog_pm] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_pm_isreaded] DEFAULT (0) FOR [isreaded],
	CONSTRAINT [DF_oblog_pm_addtime] DEFAULT (getdate()) FOR [addtime],
	CONSTRAINT [DF_oblog_pm_delR] DEFAULT (0) FOR [delR],
	CONSTRAINT [DF_oblog_pm_delS] DEFAULT (0) FOR [delS],
	CONSTRAINT [DF_oblog_pm_isguest] DEFAULT (0) FOR [isguest]
GO

ALTER TABLE [dbo].[oblog_setup] WITH NOCHECK ADD 
	CONSTRAINT [DF__oblog_set__index__2739D489] DEFAULT (0) FOR [index_to_file],
	CONSTRAINT [DF__oblog_set__df_fi__282DF8C2] DEFAULT (0) FOR [df_filename]
GO

ALTER TABLE [dbo].[oblog_subject] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_subject_subjectlognum] DEFAULT (0) FOR [subjectlognum],
	CONSTRAINT [DF_oblog_subject_ordernum] DEFAULT (0) FOR [ordernum],
	CONSTRAINT [DF__oblog_sub__subje__2B0A656D] DEFAULT (0) FOR [subjectType]
GO

ALTER TABLE [dbo].[oblog_sysskin] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_sysskin_isdefault] DEFAULT (0) FOR [isdefault]
GO

ALTER TABLE [dbo].[oblog_tags] WITH NOCHECK ADD 
	CONSTRAINT [DF__oblog_tag__iClas__19DFD96B] DEFAULT (0) FOR [iClass],
	CONSTRAINT [DF__oblog_tags__iNum__1AD3FDA4] DEFAULT (0) FOR [iNum],
	CONSTRAINT [DF__oblog_tag__iType__1BC821DD] DEFAULT (0) FOR [iType],
	CONSTRAINT [DF__oblog_tag__iStat__1CBC4616] DEFAULT (1) FOR [iState],
	CONSTRAINT [DF__oblog_tag__LastU__1DB06A4F] DEFAULT (getdate()) FOR [LastUpdate]
GO

ALTER TABLE [dbo].[oblog_trackback] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_trackback_addtime] DEFAULT (getdate()) FOR [addtime]
GO

ALTER TABLE [dbo].[oblog_upfile] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_upfile_userid] DEFAULT (0) FOR [userid],
	CONSTRAINT [DF_oblog_upfile_isphoto] DEFAULT (0) FOR [isphoto],
	CONSTRAINT [DF_oblog_upfile_addtime] DEFAULT (getdate()) FOR [addtime],
	CONSTRAINT [DF_oblog_upfile_issave] DEFAULT (0) FOR [issave],
	CONSTRAINT [DF__oblog_upf__logid__2BFE89A6] DEFAULT (0) FOR [logid],
	CONSTRAINT [DF__oblog_upf__isPow__2CF2ADDF] DEFAULT (0) FOR [isPower],
	CONSTRAINT [DF__oblog_upf__sysCl__2DE6D218] DEFAULT (0) FOR [sysClassId],
	CONSTRAINT [DF__oblog_upf__UserC__2EDAF651] DEFAULT (0) FOR [UserClassId],
	CONSTRAINT [DF__oblog_upf__photo__2FCF1A8A] DEFAULT (0) FOR [photo_width],
	CONSTRAINT [DF__oblog_upf__photo__30C33EC3] DEFAULT (0) FOR [photo_height],
	CONSTRAINT [DF__oblog_upf__viewN__31B762FC] DEFAULT (0) FOR [viewNum]
GO

ALTER TABLE [dbo].[oblog_user] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_user_sex] DEFAULT (1) FOR [sex],
	CONSTRAINT [DF_oblog_user_adddate] DEFAULT (getdate()) FOR [adddate],
	CONSTRAINT [DF_oblog_user_logintimes] DEFAULT (0) FOR [logintimes],
	CONSTRAINT [DF_oblog_user_lockuser] DEFAULT (0) FOR [lockuser],
	CONSTRAINT [DF_oblog_user_user_isbest] DEFAULT (0) FOR [user_isbest],
	CONSTRAINT [DF_oblog_user_log_count] DEFAULT (0) FOR [log_count],
	CONSTRAINT [DF_oblog_user_user_level] DEFAULT (7) FOR [user_level],
	CONSTRAINT [DF_oblog_user_user_showlog_num] DEFAULT (10) FOR [user_showlog_num],
	CONSTRAINT [DF_oblog_user_hideurl] DEFAULT (0) FOR [hideurl],
	CONSTRAINT [DF_oblog_user_user_shownewlog_num] DEFAULT (10) FOR [user_shownewlog_num],
	CONSTRAINT [DF_oblog_user_user_shownewcomment_num] DEFAULT (10) FOR [user_shownewcomment_num],
	CONSTRAINT [DF_oblog_user_user_shownewmessage_num] DEFAULT (10) FOR [user_shownewmessage_num],
	CONSTRAINT [DF_oblog_user_user_siterefu_num] DEFAULT (0) FOR [user_siterefu_num],
	CONSTRAINT [DF_oblog_user_comment_count] DEFAULT (0) FOR [comment_count],
	CONSTRAINT [DF_oblog_user_message_count] DEFAULT (0) FOR [message_count],
	CONSTRAINT [DF_oblog_user_user_upfiles_max] DEFAULT (0) FOR [user_upfiles_max],
	CONSTRAINT [DF_oblog_user_user_classid] DEFAULT (0) FOR [user_classid],
	CONSTRAINT [DF_oblog_user_user_upfiles_size] DEFAULT (0) FOR [user_upfiles_size],
	CONSTRAINT [DF_oblog_user_user_upfiles_num] DEFAULT (0) FOR [user_upfiles_num],
	CONSTRAINT [DF_oblog_user_user_showlogword_num] DEFAULT (500) FOR [user_showlogword_num],
	CONSTRAINT [DF_oblog_user_en_blogteam] DEFAULT (0) FOR [en_blogteam],
	CONSTRAINT [DF_oblog_user_comment_isasc] DEFAULT (0) FOR [comment_isasc],
	CONSTRAINT [DF_oblog_user_defaultskin] DEFAULT (0) FOR [defaultskin],
	CONSTRAINT [DF_oblog_user_isubbedit] DEFAULT (0) FOR [isubbedit],
	CONSTRAINT [DF__oblog_use__user___236943A5] DEFAULT (0) FOR [user_photorow_num]
GO

ALTER TABLE [dbo].[oblog_userclass] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_userclass_id] DEFAULT (0) FOR [id],
	CONSTRAINT [DF_oblog_userclass_classlognum] DEFAULT (0) FOR [classlognum],
	CONSTRAINT [DF_oblog_userclass_ordernum] DEFAULT (0) FOR [ordernum],
	CONSTRAINT [DF_oblog_userclass_orderid] DEFAULT (0) FOR [orderid],
	CONSTRAINT [DF_oblog_userclass_parentid] DEFAULT (0) FOR [parentid],
	CONSTRAINT [DF_oblog_userclass_depth] DEFAULT (0) FOR [depth],
	CONSTRAINT [DF_oblog_userclass_rootid] DEFAULT (0) FOR [rootid],
	CONSTRAINT [DF_oblog_userclass_child] DEFAULT (0) FOR [child],
	CONSTRAINT [DF_oblog_userclass_previd] DEFAULT (0) FOR [previd],
	CONSTRAINT [DF_oblog_userclass_nextid] DEFAULT (0) FOR [nextid],
	CONSTRAINT [DF__oblog_use__idTyp__2A164134] DEFAULT (0) FOR [idType]
GO

ALTER TABLE [dbo].[oblog_userskin] WITH NOCHECK ADD 
	CONSTRAINT [DF_oblog_userskin_isdefault] DEFAULT (0) FOR [isdefault],
	CONSTRAINT [DF_oblog_userskin_ispass] DEFAULT (0) FOR [ispass]
GO

ALTER TABLE [dbo].[oblog_usertags] WITH NOCHECK ADD 
	CONSTRAINT [DF__oblog_use__TagId__1F98B2C1] DEFAULT (0) FOR [TagId],
	CONSTRAINT [DF__oblog_use__UserI__208CD6FA] DEFAULT (0) FOR [UserId],
	CONSTRAINT [DF__oblog_use__logId__2180FB33] DEFAULT (0) FOR [logId],
	CONSTRAINT [DF__oblog_user__iNum__22751F6C] DEFAULT (1) FOR [iNum]
GO

 CREATE  INDEX [IX_oblog_admin] ON [dbo].[oblog_admin]([id], [username]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_blogstar] ON [dbo].[oblog_blogstar]([ispass]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_blogteam] ON [dbo].[oblog_blogteam]([mainuserid], [otheruserid]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_comment] ON [dbo].[oblog_comment]([mainid], [userid]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_friend] ON [dbo].[oblog_friend]([userid], [friendid], [isblack]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_log] ON [dbo].[oblog_log]([userid], [authorid], [addtime], [truetime], [classid], [subjectid], [ishide], [istop], [isbest], [isdraft]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_logclass] ON [dbo].[oblog_logclass]([id], [rootid], [ordernum]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_message] ON [dbo].[oblog_message]([userid]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_pm] ON [dbo].[oblog_pm]([sender], [incept], [isreaded], [delR], [delS]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_subject] ON [dbo].[oblog_subject]([userid], [ordernum]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_sysskin] ON [dbo].[oblog_sysskin]([isdefault]) ON [PRIMARY]
GO

 CREATE  INDEX [Index_tags_tagid] ON [dbo].[oblog_tags]([TagId]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_upfile] ON [dbo].[oblog_upfile]([userid], [isphoto]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_user] ON [dbo].[oblog_user]([username], [user_domain], [user_domainroot], [lockuser]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_userclass] ON [dbo].[oblog_userclass]([id], [ordernum], [orderid], [parentid], [rootid]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_userdir] ON [dbo].[oblog_userdir]([is_default]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_userskin] ON [dbo].[oblog_userskin]([isdefault], [ispass]) ON [PRIMARY]
GO

 CREATE  INDEX [Index_usertags_tagid] ON [dbo].[oblog_usertags]([TagId]) ON [PRIMARY]
GO

 CREATE  INDEX [Index_usertags_UserId] ON [dbo].[oblog_usertags]([UserId]) ON [PRIMARY]
GO

 CREATE  INDEX [Index_usertags_logId] ON [dbo].[oblog_usertags]([logId]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_oblog_usertags] ON [dbo].[oblog_usertags]([UserTagId]) ON [PRIMARY]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO





CREATE PROCEDURE [ob_calendar]
@logdate smalldatetime,
@userid int
 AS
SELECT addtime ,logfile
FROM oblog_log
 WHERE datediff(n,@logdate,addtime)>0 and userid=@userid



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

