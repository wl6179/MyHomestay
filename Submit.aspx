<%@ Page language="c#" Codebehind="Submit.aspx.cs" AutoEventWireup="false" Inherits="MyHomestay.Submit" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>�ύ����</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<style type="text/css"> BODY { BACKGROUND-IMAGE: none; MARGIN: 0px } </style>
		<SCRIPT type="text/JavaScript">
<!--



function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->
		</SCRIPT>
		<SCRIPT language="JavaScript">
<!--

function bt(id,after) 
{ 
eval(id+'.filters.blendTrans.stop();'); 
eval(id+'.filters.blendTrans.Apply();'); 
eval(id+'.src="'+after+'";'); 
eval(id+'.filters.blendTrans.Play();'); 
} 


//-->
		</SCRIPT>
		<!--<LINK href="css/style.css" rel="stylesheet" type="text/css">-->
		<style type="text/css"> .STYLE1 { COLOR: #666666 } .style2 { FONT-WEIGHT: bold; FONT-SIZE: 16px } </style>
		<SCRIPT language="javascript">
<!--
var where = new Array(35); 
function comefrom(loca,locacity) { this.loca = loca; this.locacity = locacity; } 
where[0]= new comefrom("��ѡ��ʡ����","��ѡ�������");
where[1] = new comefrom("����","|����|����|����|����|����|��̨|ʯ��ɽ|����|��ͷ��|��ɽ|ͨ��|˳��|��ƽ|����|ƽ��|����|����|����"); 
where[2] = new comefrom("�Ϻ�","|����|¬��|���|����|����|����|բ��|���|����|����|��ɽ|�ζ�|�ֶ�|��ɽ|�ɽ�|����|�ϻ�|����|����"); 
where[3] = new comefrom("���","|��ƽ|����|�Ӷ�|����|����|����|�Ͽ�|����|�ӱ�|����|����|����|����|���|����|����|����|����"); 
where[4] = new comefrom("����","|����|����|����|��ɿ�|����|ɳƺ��|������|�ϰ�|����|��ʢ|˫��|�山|����|ǭ��|����|�뽭|����|ͭ��|����|�ٲ�|��ɽ|��ƽ|�ǿ�|�ᶼ|�潭|��¡|����|����|����|���|��ɽ|��Ϫ|ʯ��|��ɽ|����|��ˮ|����|�ϴ�|����|�ϴ�"); 
where[5] = new comefrom("�ӱ�","|ʯ��ׯ|����|��̨|����|�żҿ�|�е�|�ȷ�|��ɽ|�ػʵ�|����|��ˮ"); 
where[6] = new comefrom("ɽ��","|̫ԭ|��ͬ|��Ȫ|����|����|˷��|����|����|����|�ٷ�|�˳�"); 
where[7] = new comefrom("���ɹ�","|���ͺ���|��ͷ|�ں�|���|���ױ�����|��������|����ľ��|�˰���|�����첼��|���ֹ�����|�����׶���|��������"); 
where[8] = new comefrom("����","|����|����|��ɽ|��˳|��Ϫ|����|����|Ӫ��|����|����|�̽�|����|����|��«��"); 
where[9] = new comefrom("����","|����|����|��ƽ|��Դ|ͨ��|��ɽ|��ԭ|�׳�|�ӱ�"); 
where[10] = new comefrom("������","|������|�������|ĵ����|��ľ˹|����|�绯|�׸�|����|�ں�|˫Ѽɽ|����|��̨��|���˰���"); 
where[11] = new comefrom("����","|�Ͼ�|��|����|��ͨ|����|�γ�|����|���Ƹ�|����|����|��Ǩ|̩��|����"); 
where[12] = new comefrom("�㽭","|����|����|����|����|����|����|��|����|��ɽ|̨��|��ˮ"); 
where[13] = new comefrom("����","|�Ϸ�|�ߺ�|����|��ɽ|����|ͭ��|����|��ɽ|����|����|����|����|����|����|����|����|����"); 
where[14] = new comefrom("����","|����|����|����|����|Ȫ��|����|��ƽ|����|����"); 
where[15] = new comefrom("����","|�ϲ���|������|�Ž�|ӥ̶|Ƽ��|����|����|����|�˴�|����|����"); 
where[16] = new comefrom("ɽ��","|����|�ൺ|�Ͳ�|��ׯ|��Ӫ|��̨|Ϋ��|����|̩��|����|����|����|����|����|�ĳ�|����|����"); 
where[17] = new comefrom("����","|֣��|����|����|ƽ��ɽ|����|�ױ�|����|����|���|���|���|����Ͽ|����|����|����|�ܿ�|פ���|��Դ"); 
where[18] = new comefrom("����","|�人|�˲�|����|�差|��ʯ|����|�Ƹ�|ʮ��|��ʩ|Ǳ��|����|����|����|����|Т��|����");
where[19] = new comefrom("����","|��ɳ|����|����|��̶|����|����|����|����|¦��|����|����|����|����|�żҽ�"); 
where[20] = new comefrom("�㶫","|����|����|�麣|��ͷ|��ݸ|��ɽ|��ɽ|�ع�|����|տ��|ï��|����|����|÷��|��β|��Դ|����|��Զ|����|����|�Ƹ�"); 
where[21] = new comefrom("����","|����|����|����|����|����|���Ǹ�|����|���|����|��������|���ݵ���|����|��ɫ|�ӳ�"); 
where[22] = new comefrom("����","|����|����"); 
where[23] = new comefrom("�Ĵ�","|�ɶ�|����|����|�Թ�|��֦��|��Ԫ|�ڽ�|��ɽ|�ϳ�|�˱�|�㰲|�ﴨ|�Ű�|üɽ|����|��ɽ|����"); 
where[24] = new comefrom("����","|����|����ˮ|����|��˳|ͭ��|ǭ����|�Ͻ�|ǭ����|ǭ��"); 
where[25] = new comefrom("����","|����|����|����|��Ϫ|��ͨ|����|���|��ɽ|˼é|��˫����|��ɽ|�º�|����|ŭ��|����|�ٲ�");
where[26] = new comefrom("����","|����|�տ���|ɽ��|��֥|����|����|����"); 
where[27] = new comefrom("����","|����|����|����|ͭ��|μ��|�Ӱ�|����|����|����|����"); 
where[28] = new comefrom("����","|����|������|���|����|��ˮ|��Ȫ|��Ҵ|����|����|¤��|ƽ��|����|����|����"); 
where[29] = new comefrom("����","|����|ʯ��ɽ|����|��ԭ"); 
where[30] = new comefrom("�ຣ","|����|����|����|����|����|����|����|����"); 
where[31] = new comefrom("�½�","|��³ľ��|ʯ����|��������|����|��������|����|�������տ¶�����|��������|��³��|����|��ʲ|����|������"); 
where[32] = new comefrom("���",""); 
where[33] = new comefrom("����",""); 
where[34] = new comefrom("̨��","|̨��|����|̨��|̨��|����|��Ͷ|����|����|�û�|����|����|����|��԰|����|��¡|̨��|����|����|���"); 
where[35] = new comefrom("����","|������|������|����|����|ŷ��|������"); 
function select() {
with(document.creator.province) { var loca2 = options[selectedIndex].value; }
for(i = 0;i < where.length;i ++) {
if (where[i].loca == loca2) {
loca3 = (where[i].locacity).split("|");
for(j = 0;j < loca3.length;j++) { with(document.creator.city) { length = loca3.length; options[j].text = loca3[j]; options[j].value = loca3[j]; var loca4=options[selectedIndex].value;}}
break;
}}
document.creator.newlocation.value=loca2+loca4;
}
function init() {
with(document.creator.province) {
length = where.length;
for(k=0;k<where.length;k++) { options[k].text = where[k].loca; options[k].value = where[k].loca; }
options[selectedIndex].text = where[0].loca; options[selectedIndex].value = where[0].loca;
}
with(document.creator.city) {
loca3 = (where[0].locacity).split("|");
length = loca3.length;
for(l=0;l<length;l++) { options[l].text = loca3[l]; options[l].value = loca3[l]; }
options[selectedIndex].text = loca3[0]; options[selectedIndex].value = loca3[0];
}}
-->
		</SCRIPT>
	</HEAD>
	<body>
		<form id="creator" method="post" runat="server">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td background="images/bg.gif" tppabs="http://www.MyHomestay.net/images/bg.gif"><div align="left"><img src="images/cloud.jpg" tppabs="http://www.MyHomestay.net/images/cloud.jpg" width="783"
								height="44"></div>
					</td>
				</tr>
			</table>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="2%" rowspan="2" valign="top" background="images/bg_left.jpg" tppabs="http://www.MyHomestay.net/images/bg_left.jpg"><img src="images/img_left.gif" tppabs="http://www.MyHomestay.net/images/img_left.gif"
							width="20" height="253"></td>
					<td width="98%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td width="2%" align="left"><img src="images/t_round.gif" tppabs="http://www.MyHomestay.net/images/t_round.gif"></td>
								<td width="98%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td width="79%">&nbsp;</td>
											<td width="13%">&nbsp;</td>
											<td width="8%">&nbsp;</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td width="19%" valign="top"><div align="center"><img src="images/logo_temp.jpg" tppabs="http://www.MyHomestay.net/images/logo_temp.jpg"
											width="142" height="53">
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td valign="top"><div align="center">
														<table width="100%" height="218" border="0" cellpadding="0" cellspacing="0">
															<tr>
																<td height="52" valign="top"><br>
																	<br>
																	<img src="images/service_Submit.jpg" width="188" height="36"><br>
																	<br>
																</td>
															</tr>
															<tr>
																<td><TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
																		<TBODY>
																			<!--<TR>
                <TD><div align="center"><A onmouseover='bt("menu1","images/menu/SubMenu01_on.jpg")' onmouseout='bt("menu1","images/menu/SubMenu01.jpg")' onfocus=this.blur() href="#" ><IMG id=menu1 style="FILTER: blendTrans(duration=0.4)" 
                        height=22 src="images/menu/SubMenu01.jpg" width=138 border=0 
                        name=menu1></A></div></TD>
              </TR>
              <TR>
                <TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.MyHomestay.net/images/menu/menu_line.jpg" width="150" height="3" /></div></TD>
              </TR>
			  
              <TR>
                <TD><div align="center"><A onmouseover='bt("menu2","images/menu/SubMenu02_on.jpg")' onmouseout='bt("menu2","images/menu/SubMenu02.jpg")' onfocus=this.blur() href="#" ><IMG id=menu2 style="FILTER: blendTrans(duration=0.4)" 
                        height=22 src="images/menu/SubMenu02.jpg" width=138 border=0 
                        name=menu2></A></div></TD>
              </TR>
              <TR>
                <TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.MyHomestay.net/images/menu/menu_line.jpg" width="150" height="3" /></div></TD>
              </TR>
			  
              <TR>
                <TD><div align="center"><A onmouseover='bt("menu3","images/menu/SubMenu03_on.jpg")' onmouseout='bt("menu3","images/menu/SubMenu03.jpg")' onfocus=this.blur() href="#" ><IMG id=menu3 style="FILTER: blendTrans(duration=0.4)" 
                        height=22 src="images/menu/SubMenu03.jpg" width=138 border=0 
                        name=menu3></A></div></TD>
              </TR>
              <TR>
                <TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.MyHomestay.net/images/menu/menu_line.jpg" width="150" height="3" /></div></TD>
              </TR>-->
																			<TR>
																				<TD background="images/menu/SubMenu00.jpg"><div style="PADDING-LEFT:43px; BACKGROUND:url(images/menu/SubMemu_FrontImage.gif) no-repeat 28px 6px"
																						align="left"><A href="SubMenu01.htm" style="TEXT-DECORATION: none"><font style="FONT-SIZE:12px;COLOR:#333333;LINE-HEIGHT:23px">ס����̼��</A></div>
																					</FONT>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.MyHomestay.net/images/menu/menu_line.jpg"
																							width="150" height="3"></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD background="images/menu/SubMenu00.jpg"><div style="PADDING-LEFT:43px; BACKGROUND:url(images/menu/SubMemu_FrontImage.gif) no-repeat 28px 6px"
																						align="left"><A href="SubMenu02.htm" style="TEXT-DECORATION: none"><font style="FONT-SIZE:12px;COLOR:#333333;LINE-HEIGHT:23px">ס���������</A></div>
																					</FONT>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.MyHomestay.net/images/menu/menu_line.jpg"
																							width="150" height="3"></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD background="images/menu/SubMenu00.jpg"><div style="PADDING-LEFT:43px; BACKGROUND:url(images/menu/SubMemu_FrontImage.gif) no-repeat 28px 6px"
																						align="left"><A href="SubMenu03.htm" style="TEXT-DECORATION: none"><font style="FONT-SIZE:12px;COLOR:#333333;LINE-HEIGHT:23px">Ϊʲôѡ��ס�����</A></div>
																					</FONT>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.MyHomestay.net/images/menu/menu_line.jpg"
																							width="150" height="3"></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD background="images/menu/SubMenu00.jpg"><div style="PADDING-LEFT:43px; BACKGROUND:url(images/menu/SubMemu_FrontImage.gif) no-repeat 28px 6px"
																						align="left"><A href="SubMenu04.htm" style="TEXT-DECORATION: none"><font style="FONT-SIZE:12px;COLOR:#333333;LINE-HEIGHT:23px">��Щ����Ҫ���</A></div>
																					</FONT>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.MyHomestay.net/images/menu/menu_line.jpg"
																							width="150" height="3"></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD background="images/menu/SubMenu00.jpg"><div style="PADDING-LEFT:43px; BACKGROUND:url(images/menu/SubMemu_FrontImage.gif) no-repeat 28px 6px"
																						align="left"><A href="service_04.htm" style="TEXT-DECORATION: none"><font style="FONT-SIZE:12px;COLOR:#333333;LINE-HEIGHT:23px">����������</A></div>
																					</FONT>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.MyHomestay.net/images/menu/menu_line.jpg"
																							width="150" height="3"></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD>&nbsp;</TD>
																			</TR>
																		</TBODY>
																	</TABLE>
																</td>
															</tr>
														</table>
													</div>
												</td>
											</tr>
											<tr>
												<td>&nbsp;</td>
											</tr>
										</table>
									</div>
								</td>
								<td width="81%" rowspan="2" valign="top"><script type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

//-->
									</script>
									<script type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
									</script>
									<table width="100%" border='0' cellpadding='0' cellspacing='0' background=" img/blank.gif"
										tppabs="http://www.MyHomestay.net/img/blank.gif">
										<tr>
											<td width='46' rowspan='2' background="images/bg_about.gif" tppabs="http://www.MyHomestay.net/images/bg_about.gif"><img src="images/top_round_1.gif" tppabs="http://www.MyHomestay.net/images/top_round_1.gif"
													width='15' height='64' alt='' border='0'></td>
											<td height='47' background="images/bg_about.gif" tppabs="http://www.MyHomestay.net/images/bg_about.gif"><a href="javascript:goUrl('about')" onMouseOver="document.mu1.src=' img/topmenu/top_mu1_01_on.gif'"
													onMouseOut="document.mu1.src=' img/topmenu/top_mu1_01_on.gif'"></a><a href="javascript:goUrl('business')" onMouseOver="document.mu2.src=' img/topmenu/top_mu1_02_on.gif'"
													onMouseOut="document.mu2.src=' img/topmenu/top_mu1_02_off.gif'"></a><a href="javascript:goUrl('ir')" onMouseOver="document.mu3.src=' img/topmenu/Menu_04_act.gif'"
													onMouseOut="document.mu3.src=' img/topmenu/Menu_03.gif'"></a><a href="javascript:goUrl('kepco_open')" onMouseOver="document.mu4.src=' img/topmenu/Menu_05_act.gif'"
													onMouseOut="document.mu4.src=' img/topmenu/Menu_05.gif'"></a>
												<table width="98%" border="0" cellspacing="0" cellpadding="0">
													<tr>
														<td width="9%"><a href="index.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image9','','images/menu/Menu_01_act.jpg'/*tpa=http://www.MyHomestay.net/images/menu/Menu_01_act.jpg*/,1)"><img src="images/menu/Menu_01.jpg" name="Image9" width="80" height="29" border="0"></a></td>
														<td width="91%"><a href="service_02.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image3','','images/menu/Menu_02_act.jpg'/*tpa=http://www.MyHomestay.net/images/menu/Menu_02_act.jpg*/,1)"><img src="images/menu/Menu_02.jpg" name="Image3" width="80" height="29" border="0"></a><a href="service_03.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image4','','images/menu/Menu_03_act.jpg'/*tpa=http://www.MyHomestay.net/images/menu/Menu_03_act.jpg*/,1)"><img src="images/menu/Menu_03.jpg" name="Image4" width="80" height="29" border="0"></a><a href="service_04.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image5','','images/menu/Menu_04_act.jpg'/*tpa=http://www.MyHomestay.net/images/menu/Menu_04_act.jpg*/,1)"><img src="images/menu/Menu_04.jpg" name="Image5" width="80" height="29" border="0"></a><a href="service_05.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image6','','images/menu/Menu_05_act.jpg'/*tpa=http://www.MyHomestay.net/images/menu/Menu_05_act.jpg*/,1)"><img src="images/menu/Menu_05.jpg" name="Image6" width="80" height="29" border="0"></a><a href="feed.aspx" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image7','','images/menu/Menu_06_act.jpg'/*tpa=http://www.MyHomestay.net/images/menu/Menu_06_act.jpg*/,1)"><img src="images/menu/Menu_06.jpg" name="Image7" width="80" height="29" border="0"></a><a href="service_07.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','images/menu/Menu_07_act.jpg'/*tpa=http://www.MyHomestay.net/images/menu/Menu_07_act.jpg*/,1)"><img src="images/menu/Menu_07.jpg" name="Image8" width="80" height="29" border="0"></a></td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td height='17' colspan='5' bgcolor='#3475ab'><img src=" img/blank.gif" tppabs="http://www.MyHomestay.net/img/blank.gif" width='1'
													height='1' alt='' border='0'></td>
										</tr>
									</table>
									<a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image3','','images/menu/Menu_02_act.jpg'/*tpa=http://www.MyHomestay.net/images/menu/Menu_02_act.jpg*/,1)">
									</a>
									<table height="100%" border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td><table cellspacing="0" cellpadding="0" width="581" border="0">
													<tbody>
														<tr>
															<td width="16" valign="top"><img height="20" alt="" src="images/round_sus.jpg" tppabs="http://www.MyHomestay.net/images/round_sus.jpg"
																	width="16" border="0"></td>
															<td class="lo" style="PADDING-RIGHT: 15px" valign="bottom" align="right" width="565"></td>
														</tr>
														<tr>
															<td colspan="2" align="left" valign="top"><img src="images/banner_submit.jpg" width="608" height="108" style="WIDTH: 608px; HEIGHT: 108px"></td>
														</tr>
														<tr>
															<td colspan="2" align="left" valign="top">&nbsp;</td>
														</tr>
														<tr>
															<td colspan="2" align="left" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
																	<tr>
																		<td colspan="2" class="int1">
																			<asp:panel id="Panel1" runat="server" Height="368px" Width="553px">
																				<TABLE id="Table1" style="WIDTH: 608px; HEIGHT: 432px" cellSpacing="1" cellPadding="1"
																					width="608" background="./images/bg1.gif" border="0" runat="server">
																					<TR>
																						<TD style="WIDTH: 17px; HEIGHT: 17px"></TD>
																						<TD style="WIDTH: 79px; HEIGHT: 17px"></TD>
																						<TD style="WIDTH: 251px; HEIGHT: 17px"></TD>
																						<TD style="HEIGHT: 17px"><FONT face="����"></FONT></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px; HEIGHT: 32px">
																							<asp:RequiredFieldValidator id="RequiredFieldValidator1" runat="server" ErrorMessage="RequiredFieldValidator"
																								ControlToValidate="txtUserName">*</asp:RequiredFieldValidator></TD>
																						<TD style="WIDTH: 79px; HEIGHT: 32px"><FONT face="����">
																								<asp:label id="lblUserName" runat="server" Width="32px" Height="18" Font-Size="X-Small">����</asp:label></FONT></TD>
																						<TD style="WIDTH: 251px; HEIGHT: 32px"><FONT face="����">
																								<asp:textbox id="txtUserName" runat="server" Width="152px" Height="24"></asp:textbox></FONT></TD>
																						<TD style="HEIGHT: 32px">
																							<asp:RegularExpressionValidator id="RegularExpressionValidator7" runat="server" ErrorMessage="RegularExpressionValidator"
																								ControlToValidate="txtUserName" Font-Size="X-Small" ValidationExpression="[\u4e00-\u9fa5]{2,6}">����д������ʵ��������</asp:RegularExpressionValidator></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px; HEIGHT: 36px"></TD>
																						<TD style="WIDTH: 79px; HEIGHT: 36px">
																							<asp:label id="lblSex" runat="server" Width="26px" Height="18px" Font-Size="X-Small">�Ա�</asp:label></TD>
																						<TD style="WIDTH: 251px; HEIGHT: 36px">
																							<asp:RadioButtonList id="RadioButtonList7" runat="server" Height="16px" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="��" Selected="True">��</asp:ListItem>
																								<asp:ListItem Value="Ů">Ů</asp:ListItem>
																							</asp:RadioButtonList></TD>
																						<TD style="HEIGHT: 36px"><FONT face="����"></FONT></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px">
																							<asp:requiredfieldvalidator id="valrEmail" runat="server" Width="8px" Height="11px" ErrorMessage="RequiredFieldValidator"
																								ControlToValidate="txtEmail">*</asp:requiredfieldvalidator></TD>
																						<TD style="WIDTH: 79px">
																							<asp:label id="lblEmail" runat="server" Font-Size="X-Small">��������</asp:label>&nbsp;&nbsp;</TD>
																						<TD style="WIDTH: 251px">
																							<asp:textbox id="txtEmail" runat="server"></asp:textbox></TD>
																						<TD>
																							<asp:regularexpressionvalidator id="valeEmail" runat="server" Width="144px" ErrorMessage="��������Ч��Email��ַ" ControlToValidate="txtEmail"
																								Font-Size="X-Small" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:regularexpressionvalidator></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px"></TD>
																						<TD style="WIDTH: 79px">
																							<asp:label id="lblBirthday" runat="server" Width="58px" Height="11px" Font-Size="X-Small">��������</asp:label></TD>
																						<TD style="WIDTH: 251px">
																							<asp:DropDownList id="ddlYear" runat="server" Width="57px">
																								<asp:ListItem Value="">ѡ��</asp:ListItem>
																								<asp:ListItem Value="1900">1900</asp:ListItem>
																								<asp:ListItem Value="1901">1901</asp:ListItem>
																								<asp:ListItem Value="1902">1902</asp:ListItem>
																								<asp:ListItem Value="1903">1903</asp:ListItem>
																								<asp:ListItem Value="1904">1904</asp:ListItem>
																								<asp:ListItem Value="1905">1905</asp:ListItem>
																								<asp:ListItem Value="1906">1906</asp:ListItem>
																								<asp:ListItem Value="1907">1907</asp:ListItem>
																								<asp:ListItem Value="1908">1908</asp:ListItem>
																								<asp:ListItem Value="1909">1909</asp:ListItem>
																								<asp:ListItem Value="1910">1910</asp:ListItem>
																								<asp:ListItem Value="1911">1911</asp:ListItem>
																								<asp:ListItem Value="1912">1912</asp:ListItem>
																								<asp:ListItem Value="1913">1913</asp:ListItem>
																								<asp:ListItem Value="1914">1914</asp:ListItem>
																								<asp:ListItem Value="1915">1915</asp:ListItem>
																								<asp:ListItem Value="1916">1916</asp:ListItem>
																								<asp:ListItem Value="1917">1917</asp:ListItem>
																								<asp:ListItem Value="1918">1918</asp:ListItem>
																								<asp:ListItem Value="1919">1919</asp:ListItem>
																								<asp:ListItem Value="1920">1920</asp:ListItem>
																								<asp:ListItem Value="1921">1921</asp:ListItem>
																								<asp:ListItem Value="1922">1922</asp:ListItem>
																								<asp:ListItem Value="1923">1923</asp:ListItem>
																								<asp:ListItem Value="1924">1924</asp:ListItem>
																								<asp:ListItem Value="1925">1925</asp:ListItem>
																								<asp:ListItem Value="1926">1926</asp:ListItem>
																								<asp:ListItem Value="1927">1927</asp:ListItem>
																								<asp:ListItem Value="1928">1928</asp:ListItem>
																								<asp:ListItem Value="1929">1929</asp:ListItem>
																								<asp:ListItem Value="1930">1930</asp:ListItem>
																								<asp:ListItem Value="1931">1931</asp:ListItem>
																								<asp:ListItem Value="1932">1932</asp:ListItem>
																								<asp:ListItem Value="1933">1933</asp:ListItem>
																								<asp:ListItem Value="1934">1934</asp:ListItem>
																								<asp:ListItem Value="1935">1935</asp:ListItem>
																								<asp:ListItem Value="1936">1936</asp:ListItem>
																								<asp:ListItem Value="1937">1937</asp:ListItem>
																								<asp:ListItem Value="1938">1938</asp:ListItem>
																								<asp:ListItem Value="1939">1939</asp:ListItem>
																								<asp:ListItem Value="1940">1940</asp:ListItem>
																								<asp:ListItem Value="1941">1941</asp:ListItem>
																								<asp:ListItem Value="1942">1942</asp:ListItem>
																								<asp:ListItem Value="1943">1943</asp:ListItem>
																								<asp:ListItem Value="1944">1944</asp:ListItem>
																								<asp:ListItem Value="1945">1945</asp:ListItem>
																								<asp:ListItem Value="1946">1946</asp:ListItem>
																								<asp:ListItem Value="1947">1947</asp:ListItem>
																								<asp:ListItem Value="1948">1948</asp:ListItem>
																								<asp:ListItem Value="1949">1949</asp:ListItem>
																								<asp:ListItem Value="1950">1950</asp:ListItem>
																								<asp:ListItem Value="1951">1951</asp:ListItem>
																								<asp:ListItem Value="1952">1952</asp:ListItem>
																								<asp:ListItem Value="1953">1953</asp:ListItem>
																								<asp:ListItem Value="1954">1954</asp:ListItem>
																								<asp:ListItem Value="1955">1955</asp:ListItem>
																								<asp:ListItem Value="1956">1956</asp:ListItem>
																								<asp:ListItem Value="1957">1957</asp:ListItem>
																								<asp:ListItem Value="1958">1958</asp:ListItem>
																								<asp:ListItem Value="1959">1959</asp:ListItem>
																								<asp:ListItem Value="1960">1960</asp:ListItem>
																								<asp:ListItem Value="1961">1961</asp:ListItem>
																								<asp:ListItem Value="1962">1962</asp:ListItem>
																								<asp:ListItem Value="1963">1963</asp:ListItem>
																								<asp:ListItem Value="1964">1964</asp:ListItem>
																								<asp:ListItem Value="1965">1965</asp:ListItem>
																								<asp:ListItem Value="1966">1966</asp:ListItem>
																								<asp:ListItem Value="1967">1967</asp:ListItem>
																								<asp:ListItem Value="1968">1968</asp:ListItem>
																								<asp:ListItem Value="1969">1969</asp:ListItem>
																								<asp:ListItem Value="1970">1970</asp:ListItem>
																								<asp:ListItem Value="1971">1971</asp:ListItem>
																								<asp:ListItem Value="1972">1972</asp:ListItem>
																								<asp:ListItem Value="1973">1973</asp:ListItem>
																								<asp:ListItem Value="1974">1974</asp:ListItem>
																								<asp:ListItem Value="1975">1975</asp:ListItem>
																								<asp:ListItem Value="1976">1976</asp:ListItem>
																								<asp:ListItem Value="1977">1977</asp:ListItem>
																								<asp:ListItem Value="1978">1978</asp:ListItem>
																								<asp:ListItem Value="1979">1979</asp:ListItem>
																								<asp:ListItem Value="1980">1980</asp:ListItem>
																								<asp:ListItem Value="1981">1981</asp:ListItem>
																								<asp:ListItem Value="1982">1982</asp:ListItem>
																								<asp:ListItem Value="1983">1983</asp:ListItem>
																								<asp:ListItem Value="1984">1984</asp:ListItem>
																								<asp:ListItem Value="1985">1985</asp:ListItem>
																								<asp:ListItem Value="1986">1986</asp:ListItem>
																								<asp:ListItem Value="1987">1987</asp:ListItem>
																								<asp:ListItem Value="1988">1988</asp:ListItem>
																								<asp:ListItem Value="1989">1989</asp:ListItem>
																								<asp:ListItem Value="1990">1990</asp:ListItem>
																								<asp:ListItem Value="1991">1991</asp:ListItem>
																								<asp:ListItem Value="1992">1992</asp:ListItem>
																								<asp:ListItem Value="1993">1993</asp:ListItem>
																								<asp:ListItem Value="1994">1994</asp:ListItem>
																								<asp:ListItem Value="1995">1995</asp:ListItem>
																								<asp:ListItem Value="1996">1996</asp:ListItem>
																								<asp:ListItem Value="1997">1997</asp:ListItem>
																								<asp:ListItem Value="1998">1998</asp:ListItem>
																								<asp:ListItem Value="1999">1999</asp:ListItem>
																								<asp:ListItem Value="2000">2000</asp:ListItem>
																								<asp:ListItem Value="2001">2001</asp:ListItem>
																								<asp:ListItem Value="2002">2002</asp:ListItem>
																								<asp:ListItem Value="2003">2003</asp:ListItem>
																								<asp:ListItem Value="2004">2004</asp:ListItem>
																								<asp:ListItem Value="2005">2005</asp:ListItem>
																								<asp:ListItem Value="2006">2006</asp:ListItem>
																							</asp:DropDownList>
																							<asp:Label id="lblYear" runat="server" Font-Size="X-Small">��</asp:Label>
																							<asp:DropDownList id="ddlMonth" runat="server">
																								<asp:ListItem Value="">ѡ��</asp:ListItem>
																								<asp:ListItem Value="1">1</asp:ListItem>
																								<asp:ListItem Value="2">2</asp:ListItem>
																								<asp:ListItem Value="3">3</asp:ListItem>
																								<asp:ListItem Value="4">4</asp:ListItem>
																								<asp:ListItem Value="5">5</asp:ListItem>
																								<asp:ListItem Value="6">6</asp:ListItem>
																								<asp:ListItem Value="7">7</asp:ListItem>
																								<asp:ListItem Value="8">8</asp:ListItem>
																								<asp:ListItem Value="9">9</asp:ListItem>
																								<asp:ListItem Value="10">10</asp:ListItem>
																								<asp:ListItem Value="11">11</asp:ListItem>
																								<asp:ListItem Value="12">12</asp:ListItem>
																							</asp:DropDownList>
																							<asp:Label id="lblMonth" runat="server" Font-Size="X-Small">��</asp:Label>
																							<asp:DropDownList id="ddlDay" runat="server" Width="56px">
																								<asp:ListItem Value="">ѡ��</asp:ListItem>
																								<asp:ListItem Value="1">1</asp:ListItem>
																								<asp:ListItem Value="2">2</asp:ListItem>
																								<asp:ListItem Value="3">3</asp:ListItem>
																								<asp:ListItem Value="4">4</asp:ListItem>
																								<asp:ListItem Value="5">5</asp:ListItem>
																								<asp:ListItem Value="6">6</asp:ListItem>
																								<asp:ListItem Value="7">7</asp:ListItem>
																								<asp:ListItem Value="8">8</asp:ListItem>
																								<asp:ListItem Value="9">9</asp:ListItem>
																								<asp:ListItem Value="10">10</asp:ListItem>
																								<asp:ListItem Value="11">11</asp:ListItem>
																								<asp:ListItem Value="12">12</asp:ListItem>
																								<asp:ListItem Value="13">13</asp:ListItem>
																								<asp:ListItem Value="14">14</asp:ListItem>
																								<asp:ListItem Value="15">15</asp:ListItem>
																								<asp:ListItem Value="16">16</asp:ListItem>
																								<asp:ListItem Value="17">17</asp:ListItem>
																								<asp:ListItem Value="18">18</asp:ListItem>
																								<asp:ListItem Value="19">19</asp:ListItem>
																								<asp:ListItem Value="20">20</asp:ListItem>
																								<asp:ListItem Value="21">21</asp:ListItem>
																								<asp:ListItem Value="22">22</asp:ListItem>
																								<asp:ListItem Value="23">23</asp:ListItem>
																								<asp:ListItem Value="24">24</asp:ListItem>
																								<asp:ListItem Value="25">25</asp:ListItem>
																								<asp:ListItem Value="26">26</asp:ListItem>
																								<asp:ListItem Value="27">27</asp:ListItem>
																								<asp:ListItem Value="28">28</asp:ListItem>
																								<asp:ListItem Value="29">29</asp:ListItem>
																								<asp:ListItem Value="30">30</asp:ListItem>
																								<asp:ListItem Value="31">31</asp:ListItem>
																							</asp:DropDownList>
																							<asp:Label id="lblDays" runat="server" Font-Size="X-Small">��</asp:Label></TD>
																						<TD>
																							<asp:RequiredFieldValidator id="RequiredFieldValidator2" runat="server" ControlToValidate="ddlYear" Font-Size="X-Small">��ѡ����</asp:RequiredFieldValidator>
																							<asp:RequiredFieldValidator id="RequiredFieldValidator3" runat="server" ControlToValidate="ddlMonth" Font-Size="X-Small">��ѡ����</asp:RequiredFieldValidator>
																							<asp:RequiredFieldValidator id="RequiredFieldValidator4" runat="server" ControlToValidate="ddlDay" Font-Size="X-Small">��ѡ����</asp:RequiredFieldValidator></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px"></TD>
																						<TD style="WIDTH: 79px">
																							<asp:Label id="Label14" runat="server" Font-Size="X-Small">ʡ��</asp:Label></TD>
																						<TD style="WIDTH: 251px"><SELECT id="Select2" style="WIDTH: 112px" onChange="select()" name="province"></SELECT><SELECT id="Select1" style="WIDTH: 104px" onChange="select()" name="city"></SELECT>
																						</TD>
																						<TD><FONT face="����">
																								<SCRIPT language="javascript">init()</SCRIPT>
																							</FONT>
																						</TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px; HEIGHT: 52px"></TD>
																						<TD style="WIDTH: 79px; HEIGHT: 52px">
																							<asp:label id="lblAddress" runat="server" Width="58px" Height="18px" Font-Size="X-Small">��ϸ��ַ</asp:label></TD>
																						<TD style="WIDTH: 251px; HEIGHT: 52px"><FONT color="#000000"><INPUT id="Text1" style="FONT-WEIGHT: bold; WIDTH: 216px; HEIGHT: 21px" type="text" size="30"
																									name="newlocation"></FONT></TD>
																						<TD><FONT face="����"></FONT></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px">
																							<asp:RequiredFieldValidator id="RequiredFieldValidator7" runat="server" ErrorMessage="RequiredFieldValidator"
																								ControlToValidate="txtMPhone">*</asp:RequiredFieldValidator></TD>
																						<TD style="WIDTH: 79px">
																							<asp:label id="lblMPhone" runat="server" Width="58px" Height="18px" Font-Size="X-Small">�ƶ��绰</asp:label>&nbsp;
																						</TD>
																						<TD style="WIDTH: 251px">
																							<asp:textbox id="txtMPhone" runat="server" Width="216px" Height="24px"></asp:textbox></TD>
																						<TD>
																							<asp:RegularExpressionValidator id="RegularExpressionValidator3" runat="server" Width="144px" Height="18px" ErrorMessage="����ȷ���������ֻ�����"
																								ControlToValidate="txtMPhone" Font-Size="X-Small" ValidationExpression="13[0-9]{9}"></asp:RegularExpressionValidator></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px"></TD>
																						<TD style="WIDTH: 79px">
																							<asp:Label id="Label4" runat="server" Font-Size="X-Small">�̶��绰</asp:Label></TD>
																						<TD style="WIDTH: 251px">
																							<asp:TextBox id="TextBox1" runat="server" Width="216px" ToolTip="����010-12345678"></asp:TextBox><FONT face="����" size="2"><BR>
																							</FONT><FONT style="FONT-SIZE: x-small" face="����">����010-85493388</FONT></TD>
																						<TD>
																							<asp:RegularExpressionValidator id="RegularExpressionValidator2" runat="server" ErrorMessage="RegularExpressionValidator"
																								ControlToValidate="TextBox1" Font-Size="X-Small" ValidationExpression="\d{3}-\d{8}|\d{4}-\d{7}">����ȷ��д�̶��绰����</asp:RegularExpressionValidator></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px; HEIGHT: 21px"></TD>
																						<TD style="WIDTH: 79px; HEIGHT: 21px">
																							<asp:Label id="Label16" runat="server" Font-Size="X-Small">��ͥ�˿�</asp:Label></TD>
																						<TD style="WIDTH: 251px; HEIGHT: 21px">
																							<asp:RadioButtonList id="RadioButtonList8" runat="server" Width="216px" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="1">1</asp:ListItem>
																								<asp:ListItem Value="2">2</asp:ListItem>
																								<asp:ListItem Value="3" Selected="True">3</asp:ListItem>
																								<asp:ListItem Value="4">4</asp:ListItem>
																								<asp:ListItem Value="����">����</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px; HEIGHT: 23px">
																							<asp:RequiredFieldValidator id="RequiredFieldValidator5" runat="server" ErrorMessage="RequiredFieldValidator"
																								ControlToValidate="TextBox4">*</asp:RequiredFieldValidator></TD>
																						<TD style="WIDTH: 79px; HEIGHT: 23px">
																							<asp:Label id="Label21" runat="server" Font-Size="X-Small">��֤��</asp:Label></TD>
																						<TD style="WIDTH: 251px; HEIGHT: 23px">
																							<asp:TextBox id="TextBox4" runat="server" Width="112px" Height="25px"></asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG style="WIDTH: 40px; HEIGHT: 20px" height="20" alt="" src="draws.aspx" width="40"></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px"></TD>
																						<TD style="WIDTH: 79px"></TD>
																						<TD style="WIDTH: 251px"><FONT face="����"></FONT></TD>
																						<TD>
																							<asp:Button id="Button3" runat="server" Width="80px" Height="26px" Font-Size="X-Small" Text="��һ��>>"
																								BorderStyle="Outset"></asp:Button></TD>
																					</TR>
																				</TABLE>
																			</asp:panel><asp:panel id="Panel2" runat="server" Height="484px" Width="552px" Visible="False">
																				<TABLE id="Table2" style="WIDTH: 584px; HEIGHT: 610px" cellSpacing="1" cellPadding="1"
																					width="584" background="./images/bg1.gif" border="1">
																					<TR>
																						<TD style="WIDTH: 120px; HEIGHT: 39px">&nbsp;
																							<asp:Label id="Label1" runat="server" Font-Size="X-Small">����Ա�</asp:Label></TD>
																						<TD style="HEIGHT: 39px">
																							<asp:RadioButtonList id="RadioButtonList9" runat="server" Width="264px" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="��" Selected="True">��</asp:ListItem>
																								<asp:ListItem Value="Ů">Ů</asp:ListItem>
																								<asp:ListItem Value="����ν">����ν</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px">&nbsp;
																							<asp:Label id="Label15" runat="server" Font-Size="X-Small">�Ƿ��˵����</asp:Label></TD>
																						<TD>
																							<asp:RadioButtonList id="RadioButtonList10" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="��">��</asp:ListItem>
																								<asp:ListItem Value="��">��</asp:ListItem>
																								<asp:ListItem Value="����ν" Selected="True">����ν</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px">&nbsp;
																							<asp:Label id="Label17" runat="server" Font-Size="X-Small">����ṹ</asp:Label></TD>
																						<TD>
																							<asp:RadioButtonList id="RadioButtonList12" runat="server" Width="312px" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="һ��һ��">һ��һ��</asp:ListItem>
																								<asp:ListItem Value="����һ��" Selected="True">����һ��</asp:ListItem>
																								<asp:ListItem Value="����һ��">����һ��</asp:ListItem>
																								<asp:ListItem Value="����">����</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px">&nbsp;
																							<asp:Label id="Label19" runat="server" Font-Size="X-Small">�Ƿ�������</asp:Label></TD>
																						<TD>
																							<asp:RadioButtonList id="RadioButtonList11" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="��" Selected="True">��</asp:ListItem>
																								<asp:ListItem Value="��">��</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px; HEIGHT: 32px">&nbsp;
																							<asp:Label id="Label2" runat="server" Font-Size="X-Small">�Ƿ��г���</asp:Label></TD>
																						<TD style="HEIGHT: 32px">
																							<asp:RadioButtonList id="RadioButtonList13" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="��" Selected="True">��</asp:ListItem>
																								<asp:ListItem Value="��">��</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px; HEIGHT: 41px">&nbsp;
																							<asp:Label id="Label3" runat="server" Font-Size="X-Small">���֪������</asp:Label></TD>
																						<TD style="HEIGHT: 41px">
																							<asp:RadioButtonList id="RadioButtonList14" runat="server" Width="312px" Height="24px" Font-Size="X-Small"
																								RepeatDirection="Horizontal">
																								<asp:ListItem Value="��ֽ��־" Selected="True">��ֽ��־</asp:ListItem>
																								<asp:ListItem Value="��վ">��վ</asp:ListItem>
																								<asp:ListItem Value="���ѽ���">���ѽ���</asp:ListItem>
																								<asp:ListItem Value="����">����</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px; HEIGHT: 39px" align="center" colSpan="2">
																							<DIV style="DISPLAY: inline; WIDTH: 152px; HEIGHT: 18px" ms_positioning="FlowLayout">�ҵ�ס������ԱЭ��</DIV>
																						</TD>
																					</TR>
																					<TR>
																						<TD style="FONT-SIZE: x-small; LINE-HEIGHT: 25px" colSpan="2">
																							<P>&nbsp;&nbsp;&nbsp;&nbsp;1. ���˱�֤�ڱ���վ�ṩ��������ʵ��׼ȷ���ҵ�ס������Ȩ����һ�б�Ҫ�ĺ�ʵ��<BR>
																								&nbsp;&nbsp;&nbsp; 2. ����ͬ���ҵ�ס���������ҵ����ϣ�����ͼƬ�ȣ������ڱ���վ�ͻ��������������վ<BR>
																								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ����,�����ҵ�ס�����͸��صĴ�������ں�������ƹ㡣<BR>
																								&nbsp;&nbsp;&nbsp; 3. �ҵ�ס������Ȩ���������Ͻ��б༭ɾ������Ȩ�������ܡ�������ȡ�����ĽӴ��ʸ�<BR>
																								&nbsp;&nbsp;&nbsp; 4. �ҵ�ס�����Ա���վ���������ս���Ȩ��<BR>
																								&nbsp;������Ϣ��֤��ʵ�������ʵ�������κκ����ȫ��������˸���<BR>
																								&nbsp;&nbsp;&nbsp;&nbsp; ���������ύ��ҳע����Ϣ������ʾ��ͬ�����ǵķ�����������Ѿ��Ķ���������ǵ���˽&nbsp; ��� 
																								δ������ͬ�⣬�ҵ�ס�������Բ���й©���ĸ�����˽��Ϣ��</P>
																						</TD>
																					</TR>
																					<TR>
																						<TD colSpan="3">
																							<P><FONT face="����">&nbsp;
																									<asp:CheckBox id="CheckBox2" runat="server" Font-Size="X-Small" Text="ͬ�Ⲣ�����ҵ�ס������ԱЭ��" Checked="True"
																										AutoPostBack="True"></asp:CheckBox>&nbsp; </FONT>
																							</P>
																							<P><FONT face="����">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																									<asp:Button id="Button4" runat="server" Width="80px" Height="26px" Font-Size="X-Small" Text="��һ��>>"
																										BorderStyle="Outset" CausesValidation="False"></asp:Button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																									<asp:Button id="Button1" runat="server" Width="80px" Height="26px" Font-Size="X-Small" Text="�ύ"
																										BorderStyle="Outset"></asp:Button></P>
																							</FONT></TD>
																					</TR>
																				</TABLE>
																			</asp:panel>
																		</td>
																	</tr>
																	<tr>
																		<td colspan="2" bordercolor="#ffffff" class="int1">
																			<p align="left">
																			</p>
																		</td>
																	</tr>
																	<tr>
																		<td colspan="2" class="int1"><p><FONT face="����"></FONT>
																			</p>
																		</td>
																	</tr>
																	<tr>
																		<td><FONT face="����"></FONT>
																		</td>
																	</tr>
																	<tr>
																		<td colspan="2" class="int1">&nbsp;</td>
																	</tr>
																</table>
															</td>
														</tr>
													</tbody>
												</table>
											</td>
											<td background="images/d_line.jpg" tppabs="http://www.MyHomestay.net/images/d_line.jpg"><img src="images/blank.gif" tppabs="http://www.MyHomestay.net/images/blank.gif" width="5"
													height="1"></td>
											<td>&nbsp;</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td valign="bottom">
									<OBJECT codeBase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0"
										height="300" width="188" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000">
										<PARAM NAME="_cx" VALUE="4974">
										<PARAM NAME="_cy" VALUE="7938">
										<PARAM NAME="FlashVars" VALUE="">
										<PARAM NAME="Movie" VALUE="images/flash/main_left.swf">
										<PARAM NAME="Src" VALUE="images/flash/main_left.swf">
										<PARAM NAME="WMode" VALUE="Window">
										<PARAM NAME="Play" VALUE="-1">
										<PARAM NAME="Loop" VALUE="-1">
										<PARAM NAME="Quality" VALUE="High">
										<PARAM NAME="SAlign" VALUE="LT">
										<PARAM NAME="Menu" VALUE="0">
										<PARAM NAME="Base" VALUE="">
										<PARAM NAME="AllowScriptAccess" VALUE="">
										<PARAM NAME="Scale" VALUE="NoScale">
										<PARAM NAME="DeviceFont" VALUE="0">
										<PARAM NAME="EmbedMovie" VALUE="0">
										<PARAM NAME="BGColor" VALUE="">
										<PARAM NAME="SWRemote" VALUE="">
										<PARAM NAME="MovieData" VALUE="">
										<PARAM NAME="SeamlessTabbing" VALUE="1">
										<PARAM NAME="Profile" VALUE="0">
										<PARAM NAME="ProfileAddress" VALUE="">
										<PARAM NAME="ProfilePort" VALUE="0">
										<PARAM NAME="AllowNetworking" VALUE="all">
										<PARAM NAME="AllowFullScreen" VALUE="false">
										<embed src="images/flash/main_left.swf" tppabs="http://www.MyHomestay.net/images/flash/main_left.swf"
											quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash"
											width="188" height="300"> </embed>
									</OBJECT>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td height="60" valign="top" background="images/bg_bottom.gif" tppabs="http://www.MyHomestay.net/images/bg_bottom.gif"><img src="images/copyright.jpg" tppabs="http://www.MyHomestay.net/images/copyright.jpg"
							width="517" height="24">
						<table width="930" border="0">
							<tr>
								<td width="37">&nbsp;</td>
								<td width="883">
									<div align="center" class="STYLE1">
										<div align="center" style="FONT-SIZE: x-small">&nbsp;�绰��&nbsp;010-85493388 
											&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�������䣺<a href="mailto:service@myhomestay.com.cn">service@myhomestay.com.cn</a>
										</div>
									</div>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>
									<div align="center" class="STYLE1">
										<div align="center" style="FONT-SIZE: x-small">&nbsp;�ͷ�QQ�� <a target="blank" style="TEXT-DECORATION: none" href="http://wpa.qq.com/msgrd?V=1&amp;Uin=401065033&amp;Site=�ҵ�ס����0081&amp;Menu=yes">
												<img src="http://wpa.qq.com/pa?p=1:401065033:1" alt="�ҵ�ס����0081" border="0" align="absMiddle"
													style="WIDTH: 74px; HEIGHT: 23px" height="23" width="74"> </a>401065033 <a style="TEXT-DECORATION: none" target="blank" href="http://wpa.qq.com/msgrd?V=1&amp;Uin=437488282&amp;Site=�ҵ�ס����0082&amp;Menu=yes">
												<img src="http://wpa.qq.com/pa?p=1:437488282:1" alt="�ҵ�ס����0082" border="0" align="absMiddle">
											</a>437488282 <a style="TEXT-DECORATION: none" target="blank" href="http://wpa.qq.com/msgrd?V=1&amp;Uin=595574668&amp;Site=�ҵ�ס����0083&amp;Menu=yes">
												<img src="http://wpa.qq.com/pa?p=1:595574668:1" alt="�ҵ�ס����0083" border="0" align="absMiddle">
											</a>595574668
										</div>
									</div>
								</td>
							</tr>
							<tr>
								<td width="37">&nbsp;</td>
								<td width="883">
									<div align="center" class="STYLE1">
										<div align="center" style="FONT-SIZE: x-small"><font color="#006499" style="FONT-SIZE: x-small">&nbsp;�ͷ�MSN��&nbsp;</font>
											<img src="images/MSN_sub.gif" width="19" height="18"><a target="blank" href="msnim:chat?contact=china200701@hotmail.com" class="index">china200701@hotmail.com</a>
											<img src="images/MSN_sub.gif" width="19" height="18"><a target="blank" href="msnim:chat?contact=china200702@hotmail.com" class="index">china200702@hotmail.com</a>
											<img src="images/MSN_sub.gif" width="19" height="18"><a target="blank" href="msnim:chat?contact=china200703@hotmail.com" class="index">china200703@hotmail.com</a>
											<div style="DISPLAY:none"><script language="javascript" type="text/javascript" src="http://js.users.51.la/990479.js"></script>
												<noscript>
													<a href="http://www.51.la/?990479" target="_blank"><img alt="��Ҫ�����ͳ��" src="http://img.users.51.la/990479.asp" style="BORDER-RIGHT:medium none; BORDER-TOP:medium none; BORDER-LEFT:medium none; BORDER-BOTTOM:medium none"></a></noscript><script language="JavaScript" type="text/javascript" src="http://c6.50bang.com/click.js?user_id=435214&amp;l=1"
													charset="gb2312"></script></div>
										</div>
									</div>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<script src="http://www.google-analytics.com/urchin.js" type="text/javascript">
			</script>
			<script type="text/javascript">
_uacct = "UA-1685551-1";
urchinTracker();
			</script>
		</form>
	</body>
</HTML>
