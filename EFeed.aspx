<%@ Page language="c#" Codebehind="EFeed.aspx.cs" AutoEventWireup="false" Inherits="MyHomestay.EFeed" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>EFeed</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<style type="text/css">BODY { BACKGROUND-IMAGE: none; MARGIN: 0px }
		</style>
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
		<link href="css/style.css" rel="stylesheet" type="text/css">
		<style type="text/css"> .STYLE1 { COLOR: #666666 } </style>
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td background="images/bg.gif" tppabs="http://www.laowaizhujia.net/images/bg.gif"><div align="left"><img src="images/cloud.jpg" tppabs="http://www.laowaizhujia.net/images/cloud.jpg" width="783"
								height="44"></div>
					</td>
				</tr>
			</table>
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td width="2%" rowspan="2" valign="top" background="images/bg_left.jpg" tppabs="http://www.laowaizhujia.net/images/bg_left.jpg"><img src="images/img_left.gif" tppabs="http://www.laowaizhujia.net/images/img_left.gif"
							width="20" height="253"></td>
					<td width="98%" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td width="2%" align="left"><img src="images/t_round.gif" tppabs="http://www.laowaizhujia.net/images/t_round.gif"></td>
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
								<td width="19%" valign="top"><div align="center"><img src="images/Eng_logo_temp.jpg" width="142" height="53">
										<table width="100%" border="0" cellspacing="0" cellpadding="0">
											<tr>
												<td valign="top"><div align="center">
														<table width="100%" height="218" border="0" cellpadding="0" cellspacing="0">
															<tr>
																<td height="52" valign="top"><br>
																	<br>
																	<img src="images/Eng_service_t16.jpg" width="188" height="36"><br>
																	<br>
																</td>
															</tr>
															<tr>
																<td><TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
																		<TBODY>
																			<TR>
																				<TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.laowaizhujia.net/images/menu/menu_line.jpg"
																							width="150" height="3"></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><A onmouseover='bt("menu2","images/menu/Eng_SubMenu002_on.jpg")' onmouseout='bt("menu2","images/menu/Eng_SubMenu002.jpg")'
																							onfocus="this.blur()" href="Eng_SubMenu01.htm"><IMG id="menu2" style="FILTER: blendTrans(duration=0.4)" height="22" src="images/menu/Eng_SubMenu002.jpg"
																								width="138" border="0" name="menu2"></A></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.laowaizhujia.net/images/menu/menu_line.jpg"
																							width="150" height="3"></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><a onMouseOver='bt("menu3","images/menu/Eng_SubMenu003_on.jpg")' onMouseOut='bt("menu3","images/menu/Eng_SubMenu003.jpg")'
																							onFocus="this.blur()" href="Eng_SubMenu02.htm"><img id="menu3" style="FILTER: blendTrans(duration=0.4)" height="22" src="images/menu/Eng_SubMenu003.jpg"
																								width="138" border="0" name="menu3"></a></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.laowaizhujia.net/images/menu/menu_line.jpg"
																							width="150" height="3"></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><A onmouseover='bt("menu4","images/menu/Eng_SubMenu004_on.jpg")' onmouseout='bt("menu4","images/menu/Eng_SubMenu004.jpg")'
																							onfocus="this.blur()" href="Eng_SubMenu03.htm"><IMG id="menu4" style="FILTER: blendTrans(duration=0.4)" height="22" src="images/menu/Eng_SubMenu004.jpg"
																								width="138" border="0" name="menu4"></A></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.laowaizhujia.net/images/menu/menu_line.jpg"
																							width="150" height="3"></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><A onmouseover='bt("menu5","images/menu/Eng_SubMenu005_on.jpg")' onmouseout='bt("menu5","images/menu/Eng_SubMenu005.jpg")'
																							onfocus="this.blur()" href="Eng_SubMenu04.htm"><IMG id="menu5" style="FILTER: blendTrans(duration=0.4)" height="22" src="images/menu/Eng_SubMenu005.jpg"
																								width="138" border="0" name="menu5"></A></div>
																				</TD>
																			</TR>
																			<TR>
																				<TD><div align="center"><img src="images/menu/menu_line.jpg" tppabs="http://www.laowaizhujia.net/images/menu/menu_line.jpg"
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
										tppabs="http://www.laowaizhujia.net/img/blank.gif">
										<tr>
											<td width='46' rowspan='2' background="images/bg_about.gif" tppabs="http://www.laowaizhujia.net/images/bg_about.gif"><img src="images/top_round_1.gif" tppabs="http://www.laowaizhujia.net/images/top_round_1.gif"
													width='15' height='64' alt='' border='0'></td>
											<td height='47' background="images/bg_about.gif" tppabs="http://www.laowaizhujia.net/images/bg_about.gif"><a href="javascript:goUrl('about')" onMouseOver="document.mu1.src=' img/topmenu/top_mu1_01_on.gif'"
													onMouseOut="document.mu1.src=' img/topmenu/top_mu1_01_on.gif'"></a><a href="javascript:goUrl('business')" onMouseOver="document.mu2.src=' img/topmenu/top_mu1_02_on.gif'"
													onMouseOut="document.mu2.src=' img/topmenu/top_mu1_02_off.gif'"></a><a href="javascript:goUrl('ir')" onMouseOver="document.mu3.src=' img/topmenu/Eng_Menu_04_act.gif'"
													onMouseOut="document.mu3.src=' img/topmenu/Eng_Menu_03.gif'"></a><a href="javascript:goUrl('kepco_open')" onMouseOver="document.mu4.src=' img/topmenu/Eng_Menu_05_act.gif'"
													onMouseOut="document.mu4.src=' img/topmenu/Eng_Menu_05.gif'"></a>
												<table width="98%" border="0" cellspacing="0" cellpadding="0">
													<tr>
														<td width="9%"><a href="Default.htm" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image9','','images/menu/Eng_Menu_01_act.jpg'/*tpa=images/menu/Eng_Menu_01_act.jpg*/,1)"><img src="images/menu/Eng_Menu_01.jpg" name="Image9" width="80" height="29" border="0"></a></td>
														<td width="91%"><a href="Eng_Service_02.htm" tppabs="http://www.laowaizhujia.net/about.asp" onMouseOut="MM_swapImgRestore()"
																onMouseOver="MM_swapImage('Image3','','images/menu/Eng_Menu_02_act.jpg'/*tpa=http://www.laowaizhujia.net/images/menu/Eng_Menu_02_act.jpg*/,1)"><img src="images/menu/Eng_Menu_02.jpg" name="Image3" width="80" height="29" border="0"></a><a href="Eng_Service_03.htm" tppabs="http://www.laowaizhujia.net/service.asp" onMouseOut="MM_swapImgRestore()"
																onMouseOver="MM_swapImage('Image4','','images/menu/Eng_Menu_03_act.jpg'/*tpa=http://www.laowaizhujia.net/images/menu/Eng_Menu_03_act.jpg*/,1)"><img src="images/menu/Eng_Menu_03.jpg" name="Image4" width="80" height="29" border="0"></a><a href="Eng_Service_04.htm" tppabs="http://www.laowaizhujia.net/Client.asp" onMouseOut="MM_swapImgRestore()"
																onMouseOver="MM_swapImage('Image5','','images/menu/Eng_Menu_04_act.jpg'/*tpa=http://www.laowaizhujia.net/images/menu/Eng_Menu_04_act.jpg*/,1)"><img src="images/menu/Eng_Menu_04.jpg" name="Image5" width="80" height="29" border="0"></a><a href="EFeed.aspx" tppabs="http://www.laowaizhujia.net/news.asp" onMouseOut="MM_swapImgRestore()"
																onMouseOver="MM_swapImage('Image6','','images/menu/Eng_Menu_05_act.jpg'/*tpa=http://www.laowaizhujia.net/images/menu/Eng_Menu_05_act.jpg*/,1)"><img src="images/menu/Eng_Menu_05.jpg" name="Image6" width="80" height="29" border="0"></a><a href="Eng_Service_06.htm" tppabs="http://www.laowaizhujia.net/rec.asp" onMouseOut="MM_swapImgRestore()"
																onMouseOver="MM_swapImage('Image7','','images/menu/Eng_Menu_06_act.jpg'/*tpa=http://www.laowaizhujia.net/images/menu/Eng_Menu_06_act.jpg*/,1)"><img src="images/menu/Eng_Menu_06.jpg" name="Image7" width="80" height="29" border="0"></a>
															<img src="images/menu/Eng_Menu_07.jpg" name="Image8" width="80" height="29" border="0"></td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td height='17' colspan='5' bgcolor='#3475ab'><img src=" img/blank.gif" tppabs="http://www.laowaizhujia.net/img/blank.gif" width='1'
													height='1' alt='' border='0'></td>
										</tr>
									</table>
									<a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image3','','images/menu/Eng_Menu_02_act.jpg'/*tpa=http://www.laowaizhujia.net/images/menu/Eng_Menu_02_act.jpg*/,1)">
									</a>
									<table height="100%" border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td><table cellspacing="0" cellpadding="0" width="581" border="0">
													<tbody>
														<tr>
															<td width="16" valign="top"><img height="20" alt="" src="images/round_sus.jpg" tppabs="http://www.laowaizhujia.net/images/round_sus.jpg"
																	width="16" border="0"></td>
															<td class="lo" style="PADDING-RIGHT: 15px" valign="bottom" align="right" width="565"></td>
														</tr>
														<tr>
															<td colspan="2" align="left" valign="top"><img src="images/Eng_banner_CONTACT.jpg" tppabs="http://www.laowaizhujia.net/images/banner_service.jpg"
																	width="581" height="108"></td>
														</tr>
														<tr>
															<td colspan="2" align="left" valign="top">&nbsp;</td>
														</tr>
														<tr>
															<td colspan="2" align="left" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
																	<tr>
																		<td colspan="2" class="Eng_int1"><p><img src="images/Eng_service_t5.jpg" width="581" height="46"></p>
																		</td>
																	</tr>
																	<tr>
																		<td colspan="2" class="Eng_int1">
																			<TABLE id="Table1" style="WIDTH: 584px; HEIGHT: 176px" cellSpacing="1" cellPadding="1"
																				width="584" border="0">
																				<TR>
																					<TD style="WIDTH: 68px"><FONT face="宋体" style="FONT-SIZE: x-small">name:</FONT></TD>
																					<TD>
																						<asp:TextBox id="TextBox1" runat="server"></asp:TextBox>
																						<asp:RequiredFieldValidator id="RequiredFieldValidator1" runat="server" ErrorMessage="RequiredFieldValidator"
																							ControlToValidate="TextBox1" Font-Size="X-Small">*</asp:RequiredFieldValidator></TD>
																				</TR>
																				<TR>
																					<TD style="WIDTH: 68px"><FONT face="宋体" style="FONT-SIZE: x-small">advice:</FONT></TD>
																					<TD style="FONT-SIZE: x-small; COLOR: #6633ff">
																						*Most inputs 100 characters<BR>
																						<asp:TextBox id="TextBox2" runat="server" Width="341px" Height="104px" TextMode="MultiLine"></asp:TextBox>
																						<asp:requiredfieldvalidator id="valrEmail" runat="server" ControlToValidate="TextBox2" ErrorMessage="RequiredFieldValidator"
																							Font-Size="X-Small" Width="8px" Height="11px">*</asp:requiredfieldvalidator>
																						<asp:CustomValidator id="CustomValidator1" runat="server" ControlToValidate="TextBox2" ErrorMessage="The number of words discovers assigns the length"
																							Font-Size="X-Small"></asp:CustomValidator></TD>
																				</TR>
																				<TR>
																					<TD style="WIDTH: 68px"></TD>
																					<TD><FONT face="宋体">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																							<asp:Button id="Button1" runat="server" Width="72px" Text="Submit" Font-Size="X-Small"></asp:Button></FONT></TD>
																				</TR>
																			</TABLE>
																		</td>
																	</tr>
																	<tr>
																		<td colspan="2" class="Eng_int1">&nbsp;</td>
																	</tr>
																</table>
															</td>
														</tr>
													</tbody>
												</table>
											</td>
											<td background="images/d_line.jpg" tppabs="http://www.laowaizhujia.net/images/d_line.jpg"><img src="images/blank.gif" tppabs="http://www.laowaizhujia.net/images/blank.gif" width="5"
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
										<embed src="images/flash/main_left.swf" tppabs="http://www.laowaizhujia.net/images/flash/main_left.swf"
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
					<td height="60" valign="top" background="images/bg_bottom.gif" tppabs="http://www.laowaizhujia.net/images/bg_bottom.gif"><img src="images/copyright.jpg" tppabs="http://www.laowaizhujia.net/images/copyright.jpg"
							width="517" height="24">
						<table width="930" border="0">
							<tr>
								<td width="37">&nbsp;</td>
								<td width="883">
									<div align="center" class="STYLE1">
										<div align="center">&nbsp;Service Tel:&nbsp;＋86-10-85493388 (Vivian) 
											&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;E-Mail:<a href="mailto:service@myhomestay.com.cn">service@myhomestay.com.cn</a>
										</div>
									</div>
								</td>
							</tr>
							<tr>
								<td width="37">&nbsp;</td>
								<td width="883">
									<div align="center" class="STYLE1">
										<div align="center"><font color="#006499">&nbsp;Service MSN:</font> <img src="images/MSN_sub.gif" width="19" height="18"><a target="blank" href="msnim:chat?contact=china200701@hotmail.com" class="index">china200701@hotmail.com</a>
											<img src="images/MSN_sub.gif" width="19" height="18"><a target="blank" href="msnim:chat?contact=china200702@hotmail.com" class="index">china200702@hotmail.com</a>
											<img src="images/MSN_sub.gif" width="19" height="18"><a target="blank" href="msnim:chat?contact=china200703@hotmail.com" class="index">china200703@hotmail.com</a>
											<div style="DISPLAY:none"><script language="javascript" type="text/javascript" src="http://js.users.51.la/990479.js"></script>
												<noscript>
													<a href="http://www.51.la/?990479" target="_blank"><img alt="我要啦免费统计" src="http://img.users.51.la/990479.asp" style="BORDER-RIGHT:medium none; BORDER-TOP:medium none; BORDER-LEFT:medium none; BORDER-BOTTOM:medium none"></a></noscript><script language="JavaScript" type="text/javascript" src="http://c6.50bang.com/click.js?user_id=435214&amp;l=1"
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
