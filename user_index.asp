<!--#include file="user_top.asp"-->


<div id="head">
	  	<div id="head2">
		
			<div id="head2-logo">
				
				<ul>
					<li class="active"><a href="http://www.myhomestay.com.cn">��������</a></li>
					<li><a href="http://www.myhomestay.com.cn">English</a></li>
					<!--<li><a href="http://www.myhomestay.com.cn">Janpan</a></li>-->
				</ul>
				
			</div>
			
			
			<div id="head2-menu">
				<div id="head2-search">
					<span>�ͻ���ѯ����</span>��<span>010-85493388</span>/<span>13910652850</span> &nbsp;&nbsp;&nbsp;
					<!--<a href="http://www.myhomestay.com.cn">��Ϊ��ҳ</a>&nbsp;&nbsp;&nbsp;-->
					<a href="#" onClick="javascript:AddFavoriteOnClick();">���ո�������ղؼ�</a>&nbsp;
					<!--<a href="http://www.myhomestay.com.cn">��������</a>&nbsp;&nbsp;&nbsp;-->
				</div>
				
				<div id="divCN-EN">
				<ul>
					<li><a href="index.html">��ҳ<br />Home</a></li>
					<li><a href="page.html">��������<br />About Us</a></li>
					<li><a href="page.html">ס�����<br />About Us</a></li>
					<li><a href="page.html">��������<br />F A Q</a></li>
					<li><a href="page.html">�ύ����<br />REGISTER</a></li>
					<li><a href="page.html">�������<br />Feedback</a></li>
					<li><a href="page.html">��ϵ����<br />Contact Us</a></li>
					<li class="active">��ѯ��Ϣ<br />consultation</li>

					
				</ul>
				</div>
			</div>
		
		</div>
		
	  </div>
	  
	  
	  <div id="body"><!--#body-->
	  	<div id="gbody"><!--#gbody-->
		
			<div id="left-margin"></div>
			
			<div id="main"><!--the main content area-->
				
				
				
				
				<div id="mainContentList">
					<div id="main-top-round">
						<div id="main-top-round-left"></div>
						<div id="main-top-round-right"></div>
					</div>
					
					<div id="main-main-round">
						<style>
							table {
							background:#ffffff;
							border:1px solid #999999;
							border-collapse: collapse;
							font-size:13px;
							}
							td {
							padding:0px 3px 0px 12px;
							width:100px;
							height:18px;
							border:1px solid #999999;
							}
							.tr01 {
							/*font:bold;*/
							background:#A1E5FA;
							font-boil:2px;
							color:#4f6b72;
							}
							.tr02 {
							background:#ffffff;
							color:#797268;
							}
							.tr03 {
							background:#F7FDFB;
							color:#4f6b72;
							}
						</style>
						<div id="main-TextArea01">
						
						
						
							<center><h2>�����Ů����ס��</h2></center>
							
							
							<%main()%>
				
						
					<!--wl-->
					</div>
					
					
					<div id="main-main-round2">
						<center>
					  
					    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="585" height="90">
                          <param name="movie" value="images/90620_hp585_90.swf" />
                          <param name="quality" value="high" />
                          <embed src="images/90620_hp585_90.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="585" height="90"></embed>
					      </object>
					  	</center>
					</div>
					
					
		
					
					
				
					<div id="main-banner01">
						<div id="main-bottom-round-left"></div>
						<div id="main-bottom-round-right"></div>
					</div>
				</div>
				
				
				<div id="publicList">
					<div id="public-top-round">
						<div id="public-top-round-left"></div>
						<div id="public-top-round-right"></div>
					</div>
					
					<div id="public-main-round">
					<%'main()%>
						<div class="public-TextArea01">
							<p id="public-title01">
								��Ŀ��������
							</p>
							
							<!--#include file="menu/inc-menu.asp"-->
							<p>
								<ul>
									<li><a href="back-ctrl.html">��һ��homestay�</a></li>
									<li><a href="back-ctrl.html">�����Ů����ס��</a></li>
									<li><a href="back-ctrl.html">����20���к���ס</a></li>
									<li><a href="back-ctrl.html">AMERICAN MOSAIC</a></li>
									<li><a href="back-ctrl.html">[����sohu�] ��</a></li>
									<li><a href="back-ctrl.html">����Ӣ��100���</a></li>
									<li><a href="back-ctrl.html">��,����һ����</a></li>
									<li><a href="back-ctrl.html">���˶��ܿ�ѷ����</a></li>
								</ul>
							</p>
						</div>
						
					
					</div>
				
					<div id="public-bottom-round">
						<div id="public-bottom-round-left"></div>
						<div id="public-bottom-round-right"></div>
					</div>
				</div>				
			</div>
			

		</div><!--#gbody-->	
	  </div><!--#body-->
	  
	  
	  <div id="bottom"><!--bottom-->
	  	<div id="bottom2"><!--bottom2-->
	  		
<center><%=oblog.user_copyright%></center>
			
			
	  	</div><!--bottom2-->
	  </div><!--bottom-->
  
  <script>
  	//��A�ͻ�����ղ�-��ת����ͼ��Ƭ����ҳ,�밴A...97
	//www.myhomestay.com.cn  wangliang6179@163.com
	var hotkey=32
	var destination="http://www.myhomestay.com.cn"
	if (document.layers)
	document.captureEvents(Event.KEYPRESS)
	function backhome(e){
		if (document.layers){
		if (e.which==hotkey)
		window.external.AddFavorite(document.location.href,document.title)
		//window.location=destination
		//alert("лл����עHomestay�����Ļ����� �ҵ�ס������������ѯ���ߣ�010-85493388");
		}
		else if (document.all){
		if (event.keyCode==hotkey)
		window.external.AddFavorite(document.location.href,document.title);
		//alert("лл����עHomestay�����Ļ����� �ҵ�ס������������ѯ���ߣ�010-85493388");
		//window.location=destination
		}
	}
	document.onkeypress=backhome
	onkeydown="javascript:onenter();"
	
	function onenter(){
		 if(event.keyCode==13){
		//alert("лл����עHomestay�����Ļ����� �ҵ�ס������������ѯ���ߣ�010-85493388");
		}
	}
	
	function AddFavoriteOnClick(){
		window.external.AddFavorite(document.location.href,document.title)
		//alert("лл����עHomestay�����Ļ����� �ҵ�ס������������ѯ���ߣ�010-85493388");
	}
  </script>
  <script src="inc/divCN-EN.js" type="text/javascript"></script>
  </body>
</html>
<%
sub main
dim rs
%>
<div class="msg_left">
<div class="msg">
	<h1>�ҵ�δ������</h1>
	<ul>
<%
set rs=oblog.execute("select id,sender,topic,addtime,content from oblog_pm where incept='"&oblog.logined_uname&"' and delr=0 and isreaded=0 order by id desc")
while not rs.eof
	response.Write("<li><span class=""msg_user"">"&oblog.filt_html(rs("sender"))&"��</span><a href=""javascript:openScript('user_pm.asp?action=read&id="&rs("id")&"',450,380)"" title="""&oblog.filt_html(rs("content"))&""">"&oblog.filt_html(rs("topic"))&"</a> <span class=""msg_date"">"&rs("addtime")&"</span></li>")
	rs.movenext
wend
%>
	</ul>
</div>
<div class="msg">
	<h1>�ҵ����10������</h1>
	<ul>
<%
set rs=oblog.execute("select top 10 commenttopic,addtime,commentid,comment_user,mainid,comment from oblog_comment where userid="&oblog.logined_uid&" order by commentid desc")
while not rs.eof
	response.Write("<li><span class=""msg_user"">"&rs("comment_user")&"��</span><a href='go.asp?logid="&rs("mainid")&"&commentid="&rs("commentid")&"' target='_blank' title="""&oblog.filt_html(left(rs("comment"),150))&""">"&rs("commenttopic")&"</a> <span class=""msg_date"">"&rs("addtime")&"</span></li>")
	rs.movenext
wend
%>
	</ul>
</div>
<div class="msg">
	<h1>�ҵ����10������</h1>
	<ul>
<%
set rs=oblog.execute("select top 10 messagetopic,addtime,messageid,message_user,message from oblog_message where userid="&oblog.logined_uid&" order by messageid desc")
while not rs.eof
	response.Write("<li><span class=""msg_user"">"&rs("message_user")&"��</span><a href='go.asp?messageid="&rs("messageid")&"' target='_blank' title="""&oblog.filt_html(left(rs("message"),150))&""">"&rs("messagetopic")&"</a> <span class=""msg_date"">"&rs("addtime")&"</span></li>")
	rs.movenext
wend
set rs=nothing
%>
	</ul>
</div>
</div>

<div class="msg_site">

	<h5>��ӭ�û�</h5>
	<p><%=oblog.setup(20,0)%></p>

</div>

<%end sub%>