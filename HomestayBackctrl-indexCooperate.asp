<!--#include file="user_topCooperate.asp"-->

    
  
  
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
					<span id="joincompany_login"><a href="/login.asp">��ҵ��¼</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span id="joincompany"><a href="/RegisterCooperate.asp">��ҵ����</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span>�ͻ���ѯ����</span>��<span>010-85493388</span>/<span>13146398085</span> &nbsp;&nbsp;&nbsp;
					<!--<a href="http://www.myhomestay.com.cn">��Ϊ��ҳ</a>&nbsp;&nbsp;&nbsp;-->
					<!--<a href="#" onClick="javascript:AddFavoriteOnClick();">���ո�������ղؼ�</a>-->&nbsp;
					<!--<a href="http://www.myhomestay.com.cn">��������</a>&nbsp;&nbsp;&nbsp;-->
				</div>
				
				<div id="divCN-EN">
				<ul>
					<li><a href="/index.asp">��ҳ<br />Home</a></li><!--#include file="menu/incIndexNav.htm"-->

					
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
							border:1px solid #FFFFFF;
							border-collapse: collapse;
							font-size:13px;
							}
							table#tableInfo {
							float:left;
							}
							
							div#divPhoto {
							float:left;
							border:1px #999999 solid;
							}
							div#divPhoto img {
								border:1px # solid;
							}
							
							td {
							padding:0px 3px 0px 12px;
							/*width:100px;*/
							height:18px;
							border:1px solid #FFFFFF;
							}
							.tr01 {
							/*font:bold;*/
							/*background:#A1E5FA;*/
							background:#ffffff;
							font-boil:2px;
							color:#4f6b72;
							}
							.tr01 .tdMap {
							width:180px;
							}
							.tr01 .tdColspan2 {
							width:330px;
							}
							
							.tr02 {
							background:#ffffff;
							color:#797268;
							}
							.tr03 {
							background:#ffffff;
							color:#797268;
							}
						</style>
						<div id="main-TextArea01">
						
						
						
							<center><h2>���˹���������ҳ</h2></center>
							
<%
dim rs, sql,action
dim id,usersearch,Keyword,strField

usersearch=trim(request("usersearch"))
action=trim(request("action"))
id=oblog.filt_badstr(trim(Request("id")))
if usersearch="" then
	usersearch=0
else
	usersearch=Clng(usersearch)
end if


select case action
	case "modifyphoto"
	call modify()
	case "savemodify"
	call Savemodify()
	case "delfile" 
	call delfile()

	case else
	call main2()
	
end select
set rs=nothing
%>							
							

						</div>					
						
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
					<%main()%>
						<div class="public-TextArea01">
							<!--<p class="public-title01">
								��Ŀ��������
							</p>-->
							
							<!--#include file="menu/inc-menuCooperate.asp"-->
							
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
	  
	  
<%=oblog.user_copyright%>
  

  <script src="inc/divCN-EN.js" type="text/javascript"></script>
  </body>
</html>


<%
sub main

session("indexCooperateLogined")="indexCooperateLogined"
dim rs
%>

<div class="msg_site">

	<h5>��ӭ�û�</h5>
	<p><%=oblog.setup(20,0)%></p>

</div>

<%end sub%>


<%
sub main2
dim rs,sql
dim ssql
	ssql="*"

	sql="select "&ssql&" from [oblog_user] where userid="& oblog.logined_uid &" "

Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,Conn,1,1
  	if rs.eof and rs.bof then
		Response.Write "<input type='button' name='historybackwl' value='������һҳ' onclick='javascript:history.go(-1);' class=btxx style='cursor:hand;'>"
	else
		
	end if
%>

	<table id="tableInfo">
	
		<tr class="tr01">
			<td colspan="2" rowspan="7" class="tdMap">
	<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="150" height="150">
	<param name=movie value="images/chart.swf?in6=<%=int((3/100)*100)%>&in6=<%
		Dim creditTemp'���ö�
		If rs("passTeacher_count")= 0 Then
			creditTemp= 0
		ElseIf Cint(rs("passTeacher_count"))<>0 And rs("passTeacher_count")<=100 Then
			'''creditTemp= (rs("passTeacher_count")/rs("log_count"))*100
			creditTemp= (rs("passTeacher_count")/100)*100
		ElseIf Cint(rs("passTeacher_count"))<>0 And rs("passTeacher_count")>100 And rs("passTeacher_count")<=200 Then
			'''creditTemp= 0
			creditTemp= (rs("passTeacher_count")/200)*100
		ElseIf Cint(rs("passTeacher_count"))<>0 And rs("passTeacher_count")>200 Then'�߼����к������
			'''creditTemp= 0
			creditTemp= (200/200)*100
		End If
		
		Response.Write creditTemp
	%>">
	<param name=quality value=high>
	<param name="wmode" value="transparent">
	<embed src="images/chart.swf?in6=<%=int((8/100)*100)%>&in6=<%=int(((100-92)/100)*100)%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="150" height="150">				</embed>
	</object>
	<br /><center>
	<%
	If rs("isValidate")=1 Then
	
	%>
		<% If rs("passTeacher_count")>0 Then %>
		��������
		<% End If %>
		<% If rs("passTeacher_count")>=100 Then %>
		+ �м�����
		<% End If %>
		<% If rs("passTeacher_count")>=200 Then %>
		<br />+ �߼�����
		<% End If %>
		���öȣ�<%=Round(creditTemp, 2)%>%</center>
		
		<center>
		<% If rs("passTeacher_count")>0 Then %>
		<img src="/images/MyHomestayMedal_02.gif" title="����������" />
		<% End If %>
		<% If rs("passTeacher_count")>=100 Then %>
		<img src="/images/MyHomestayMedal_02.gif" title="��Ϊ��������������" />
		<% End If %>
		<% If rs("passTeacher_count")>=200 Then %>
		<img src="/images/MyHomestayMedal_02.gif" title="���ѳ�Ϊ���������к����ĸ߼�������" />
		<% End If %>
	<%	
	ElseIf rs("isValidate")=0 Then
		'''response.Write("�ȴ������...")
	End If
	%>
	</center>
	</td>
			<td colspan="2" class="tdColspan2">��ӭ���˹����ˣ�<%=rs("blogname")%>
			</td>
		</tr>
		
		<tr class="tr02">
			<td>��˾����:</td>
			<td>
			<%			
			
	 response.Write rs("companyname")
			%> </td>
		</tr>
		<tr class="tr02">
			<td>���ڵ���(ʡ/��):</td>
			<td><%=rs("province")%>&nbsp;<%=rs("city")%></td>
		</tr>
		<tr class="tr02">
			<td>��ǰ���״̬:</td>
			<td><%
			If rs("isValidate")=1 Then
				response.Write("<font color='#038ad7'>��</font>��ͨ�����")
			ElseIf rs("isValidate")=0 Then
				response.Write("<font color=red>�ȴ������...</font>")
			End If
			%></td>
		</tr>
		
		
		
		<tr class="tr02">
			<td>����˵��������:</td>
			<td><font color='#038ad7'>��</font>�ɹ���� <font color="#038ad7"><%=rs("passTeacher_count")%></font> ��</td>
		</tr>
		<tr class="tr02">
			<td>���ύ���������:</td>
			<td>�ܹ� <font color="red"><%=rs("log_count")%></font> ���������</td>
		</tr>
		
		<tr class="tr03">
			<td>�ϴε�¼ʱ��</td>
			<td><%=rs("lastlogintime")%></td>
		</tr>
		
	</table>
								
	<div id="divPhoto" style=" position:absolute;background:url('/images/homestayCooperateSign2<% If rs("isValidate")=0 Then Response.Write("_wait")%>.gif') left top;width:170px;height:120px;padding:3px;border:0px;">
	<!--<% If rs("upfile_isCoverPhoto_path")<>"" And Len(Trim(rs("upfile_isCoverPhoto_path")))>6 Then %>
		<a target="_blank" href="<%=rs("upfile_isCoverPhoto_path")%>">
		
		<img src="<%=rs("upfile_isCoverPhoto_path")%>" style='width:110px;height:130px;padding:3px;' />
		
		</a>
	<% Else %>-->
		<!--<a href="HomestayBackctrl-PhotofilesHome.asp">-->
		<!--<img src="images/homestay.gif" style='width:110px;height:130px;padding:3px;' />-->
		<!--</a>
	<% End If %>-->
		
		<center>
		&nbsp;<span style="font-size:8px;color:#666666;"></span>
		</center>
	</div>
							
<%
rs.close
set rs=nothing

end sub



%>