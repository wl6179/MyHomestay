<%@ Page language="c#" Codebehind="Submit.aspx.cs" AutoEventWireup="false" Inherits="MyHomestay.Submit" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>提交申请</title>
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
where[0]= new comefrom("请选择省份名","请选择城市名");
where[1] = new comefrom("北京","|东城|西城|崇文|宣武|朝阳|丰台|石景山|海淀|门头沟|房山|通州|顺义|昌平|大兴|平谷|怀柔|密云|延庆"); 
where[2] = new comefrom("上海","|黄浦|卢湾|徐汇|长宁|静安|普陀|闸北|虹口|杨浦|闵行|宝山|嘉定|浦东|金山|松江|青浦|南汇|奉贤|崇明"); 
where[3] = new comefrom("天津","|和平|东丽|河东|西青|河西|津南|南开|北辰|河北|武清|红挢|塘沽|汉沽|大港|宁河|静海|宝坻|蓟县"); 
where[4] = new comefrom("重庆","|万州|涪陵|渝中|大渡口|江北|沙坪坝|九龙坡|南岸|北碚|万盛|双挢|渝北|巴南|黔江|长寿|綦江|潼南|铜梁|大足|荣昌|壁山|梁平|城口|丰都|垫江|武隆|忠县|开县|云阳|奉节|巫山|巫溪|石柱|秀山|酉阳|彭水|江津|合川|永川|南川"); 
where[5] = new comefrom("河北","|石家庄|邯郸|邢台|保定|张家口|承德|廊坊|唐山|秦皇岛|沧州|衡水"); 
where[6] = new comefrom("山西","|太原|大同|阳泉|长治|晋城|朔州|吕梁|忻州|晋中|临汾|运城"); 
where[7] = new comefrom("内蒙古","|呼和浩特|包头|乌海|赤峰|呼伦贝尔盟|阿拉善盟|哲里木盟|兴安盟|乌兰察布盟|锡林郭勒盟|巴彦淖尔盟|伊克昭盟"); 
where[8] = new comefrom("辽宁","|沈阳|大连|鞍山|抚顺|本溪|丹东|锦州|营口|阜新|辽阳|盘锦|铁岭|朝阳|葫芦岛"); 
where[9] = new comefrom("吉林","|长春|吉林|四平|辽源|通化|白山|松原|白城|延边"); 
where[10] = new comefrom("黑龙江","|哈尔滨|齐齐哈尔|牡丹江|佳木斯|大庆|绥化|鹤岗|鸡西|黑河|双鸭山|伊春|七台河|大兴安岭"); 
where[11] = new comefrom("江苏","|南京|镇江|苏州|南通|扬州|盐城|徐州|连云港|常州|无锡|宿迁|泰州|淮安"); 
where[12] = new comefrom("浙江","|杭州|宁波|温州|嘉兴|湖州|绍兴|金华|衢州|舟山|台州|丽水"); 
where[13] = new comefrom("安徽","|合肥|芜湖|蚌埠|马鞍山|淮北|铜陵|安庆|黄山|滁州|宿州|池州|淮南|巢湖|阜阳|六安|宣城|亳州"); 
where[14] = new comefrom("福建","|福州|厦门|莆田|三明|泉州|漳州|南平|龙岩|宁德"); 
where[15] = new comefrom("江西","|南昌市|景德镇|九江|鹰潭|萍乡|新馀|赣州|吉安|宜春|抚州|上饶"); 
where[16] = new comefrom("山东","|济南|青岛|淄博|枣庄|东营|烟台|潍坊|济宁|泰安|威海|日照|莱芜|临沂|德州|聊城|滨州|菏泽"); 
where[17] = new comefrom("河南","|郑州|开封|洛阳|平顶山|安阳|鹤壁|新乡|焦作|濮阳|许昌|漯河|三门峡|南阳|商丘|信阳|周口|驻马店|济源"); 
where[18] = new comefrom("湖北","|武汉|宜昌|荆州|襄樊|黄石|荆门|黄冈|十堰|恩施|潜江|天门|仙桃|随州|咸宁|孝感|鄂州");
where[19] = new comefrom("湖南","|长沙|常德|株洲|湘潭|衡阳|岳阳|邵阳|益阳|娄底|怀化|郴州|永州|湘西|张家界"); 
where[20] = new comefrom("广东","|广州|深圳|珠海|汕头|东莞|中山|佛山|韶关|江门|湛江|茂名|肇庆|惠州|梅州|汕尾|河源|阳江|清远|潮州|揭阳|云浮"); 
where[21] = new comefrom("广西","|南宁|柳州|桂林|梧州|北海|防城港|钦州|贵港|玉林|南宁地区|柳州地区|贺州|百色|河池"); 
where[22] = new comefrom("海南","|海口|三亚"); 
where[23] = new comefrom("四川","|成都|绵阳|德阳|自贡|攀枝花|广元|内江|乐山|南充|宜宾|广安|达川|雅安|眉山|甘孜|凉山|泸州"); 
where[24] = new comefrom("贵州","|贵阳|六盘水|遵义|安顺|铜仁|黔西南|毕节|黔东南|黔南"); 
where[25] = new comefrom("云南","|昆明|大理|曲靖|玉溪|昭通|楚雄|红河|文山|思茅|西双版纳|保山|德宏|丽江|怒江|迪庆|临沧");
where[26] = new comefrom("西藏","|拉萨|日喀则|山南|林芝|昌都|阿里|那曲"); 
where[27] = new comefrom("陕西","|西安|宝鸡|咸阳|铜川|渭南|延安|榆林|汉中|安康|商洛"); 
where[28] = new comefrom("甘肃","|兰州|嘉峪关|金昌|白银|天水|酒泉|张掖|武威|定西|陇南|平凉|庆阳|临夏|甘南"); 
where[29] = new comefrom("宁夏","|银川|石嘴山|吴忠|固原"); 
where[30] = new comefrom("青海","|西宁|海东|海南|海北|黄南|玉树|果洛|海西"); 
where[31] = new comefrom("新疆","|乌鲁木齐|石河子|克拉玛依|伊犁|巴音郭勒|昌吉|克孜勒苏柯尔克孜|博尔塔拉|吐鲁番|哈密|喀什|和田|阿克苏"); 
where[32] = new comefrom("香港",""); 
where[33] = new comefrom("澳门",""); 
where[34] = new comefrom("台湾","|台北|高雄|台中|台南|屏东|南投|云林|新竹|彰化|苗栗|嘉义|花莲|桃园|宜兰|基隆|台东|金门|马祖|澎湖"); 
where[35] = new comefrom("其它","|北美洲|南美洲|亚洲|非洲|欧洲|大洋洲"); 
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
																						align="left"><A href="SubMenu01.htm" style="TEXT-DECORATION: none"><font style="FONT-SIZE:12px;COLOR:#333333;LINE-HEIGHT:23px">住家外教简介</A></div>
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
																						align="left"><A href="SubMenu02.htm" style="TEXT-DECORATION: none"><font style="FONT-SIZE:12px;COLOR:#333333;LINE-HEIGHT:23px">住家外教优势</A></div>
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
																						align="left"><A href="SubMenu03.htm" style="TEXT-DECORATION: none"><font style="FONT-SIZE:12px;COLOR:#333333;LINE-HEIGHT:23px">为什么选择住家外教</A></div>
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
																						align="left"><A href="SubMenu04.htm" style="TEXT-DECORATION: none"><font style="FONT-SIZE:12px;COLOR:#333333;LINE-HEIGHT:23px">那些人需要外教</A></div>
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
																						align="left"><A href="service_04.htm" style="TEXT-DECORATION: none"><font style="FONT-SIZE:12px;COLOR:#333333;LINE-HEIGHT:23px">→现在申请</A></div>
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
																						<TD style="HEIGHT: 17px"><FONT face="宋体"></FONT></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px; HEIGHT: 32px">
																							<asp:RequiredFieldValidator id="RequiredFieldValidator1" runat="server" ErrorMessage="RequiredFieldValidator"
																								ControlToValidate="txtUserName">*</asp:RequiredFieldValidator></TD>
																						<TD style="WIDTH: 79px; HEIGHT: 32px"><FONT face="宋体">
																								<asp:label id="lblUserName" runat="server" Width="32px" Height="18" Font-Size="X-Small">姓名</asp:label></FONT></TD>
																						<TD style="WIDTH: 251px; HEIGHT: 32px"><FONT face="宋体">
																								<asp:textbox id="txtUserName" runat="server" Width="152px" Height="24"></asp:textbox></FONT></TD>
																						<TD style="HEIGHT: 32px">
																							<asp:RegularExpressionValidator id="RegularExpressionValidator7" runat="server" ErrorMessage="RegularExpressionValidator"
																								ControlToValidate="txtUserName" Font-Size="X-Small" ValidationExpression="[\u4e00-\u9fa5]{2,6}">请填写您的真实中文姓名</asp:RegularExpressionValidator></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px; HEIGHT: 36px"></TD>
																						<TD style="WIDTH: 79px; HEIGHT: 36px">
																							<asp:label id="lblSex" runat="server" Width="26px" Height="18px" Font-Size="X-Small">性别</asp:label></TD>
																						<TD style="WIDTH: 251px; HEIGHT: 36px">
																							<asp:RadioButtonList id="RadioButtonList7" runat="server" Height="16px" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="男" Selected="True">男</asp:ListItem>
																								<asp:ListItem Value="女">女</asp:ListItem>
																							</asp:RadioButtonList></TD>
																						<TD style="HEIGHT: 36px"><FONT face="宋体"></FONT></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px">
																							<asp:requiredfieldvalidator id="valrEmail" runat="server" Width="8px" Height="11px" ErrorMessage="RequiredFieldValidator"
																								ControlToValidate="txtEmail">*</asp:requiredfieldvalidator></TD>
																						<TD style="WIDTH: 79px">
																							<asp:label id="lblEmail" runat="server" Font-Size="X-Small">电子邮箱</asp:label>&nbsp;&nbsp;</TD>
																						<TD style="WIDTH: 251px">
																							<asp:textbox id="txtEmail" runat="server"></asp:textbox></TD>
																						<TD>
																							<asp:regularexpressionvalidator id="valeEmail" runat="server" Width="144px" ErrorMessage="请输入有效的Email地址" ControlToValidate="txtEmail"
																								Font-Size="X-Small" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:regularexpressionvalidator></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px"></TD>
																						<TD style="WIDTH: 79px">
																							<asp:label id="lblBirthday" runat="server" Width="58px" Height="11px" Font-Size="X-Small">您的生日</asp:label></TD>
																						<TD style="WIDTH: 251px">
																							<asp:DropDownList id="ddlYear" runat="server" Width="57px">
																								<asp:ListItem Value="">选择</asp:ListItem>
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
																							<asp:Label id="lblYear" runat="server" Font-Size="X-Small">年</asp:Label>
																							<asp:DropDownList id="ddlMonth" runat="server">
																								<asp:ListItem Value="">选择</asp:ListItem>
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
																							<asp:Label id="lblMonth" runat="server" Font-Size="X-Small">月</asp:Label>
																							<asp:DropDownList id="ddlDay" runat="server" Width="56px">
																								<asp:ListItem Value="">选择</asp:ListItem>
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
																							<asp:Label id="lblDays" runat="server" Font-Size="X-Small">日</asp:Label></TD>
																						<TD>
																							<asp:RequiredFieldValidator id="RequiredFieldValidator2" runat="server" ControlToValidate="ddlYear" Font-Size="X-Small">请选择年</asp:RequiredFieldValidator>
																							<asp:RequiredFieldValidator id="RequiredFieldValidator3" runat="server" ControlToValidate="ddlMonth" Font-Size="X-Small">请选择月</asp:RequiredFieldValidator>
																							<asp:RequiredFieldValidator id="RequiredFieldValidator4" runat="server" ControlToValidate="ddlDay" Font-Size="X-Small">请选择日</asp:RequiredFieldValidator></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px"></TD>
																						<TD style="WIDTH: 79px">
																							<asp:Label id="Label14" runat="server" Font-Size="X-Small">省份</asp:Label></TD>
																						<TD style="WIDTH: 251px"><SELECT id="Select2" style="WIDTH: 112px" onChange="select()" name="province"></SELECT><SELECT id="Select1" style="WIDTH: 104px" onChange="select()" name="city"></SELECT>
																						</TD>
																						<TD><FONT face="宋体">
																								<SCRIPT language="javascript">init()</SCRIPT>
																							</FONT>
																						</TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px; HEIGHT: 52px"></TD>
																						<TD style="WIDTH: 79px; HEIGHT: 52px">
																							<asp:label id="lblAddress" runat="server" Width="58px" Height="18px" Font-Size="X-Small">详细地址</asp:label></TD>
																						<TD style="WIDTH: 251px; HEIGHT: 52px"><FONT color="#000000"><INPUT id="Text1" style="FONT-WEIGHT: bold; WIDTH: 216px; HEIGHT: 21px" type="text" size="30"
																									name="newlocation"></FONT></TD>
																						<TD><FONT face="宋体"></FONT></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px">
																							<asp:RequiredFieldValidator id="RequiredFieldValidator7" runat="server" ErrorMessage="RequiredFieldValidator"
																								ControlToValidate="txtMPhone">*</asp:RequiredFieldValidator></TD>
																						<TD style="WIDTH: 79px">
																							<asp:label id="lblMPhone" runat="server" Width="58px" Height="18px" Font-Size="X-Small">移动电话</asp:label>&nbsp;
																						</TD>
																						<TD style="WIDTH: 251px">
																							<asp:textbox id="txtMPhone" runat="server" Width="216px" Height="24px"></asp:textbox></TD>
																						<TD>
																							<asp:RegularExpressionValidator id="RegularExpressionValidator3" runat="server" Width="144px" Height="18px" ErrorMessage="请正确输入您的手机号码"
																								ControlToValidate="txtMPhone" Font-Size="X-Small" ValidationExpression="13[0-9]{9}"></asp:RegularExpressionValidator></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px"></TD>
																						<TD style="WIDTH: 79px">
																							<asp:Label id="Label4" runat="server" Font-Size="X-Small">固定电话</asp:Label></TD>
																						<TD style="WIDTH: 251px">
																							<asp:TextBox id="TextBox1" runat="server" Width="216px" ToolTip="例：010-12345678"></asp:TextBox><FONT face="宋体" size="2"><BR>
																							</FONT><FONT style="FONT-SIZE: x-small" face="宋体">例：010-85493388</FONT></TD>
																						<TD>
																							<asp:RegularExpressionValidator id="RegularExpressionValidator2" runat="server" ErrorMessage="RegularExpressionValidator"
																								ControlToValidate="TextBox1" Font-Size="X-Small" ValidationExpression="\d{3}-\d{8}|\d{4}-\d{7}">请正确填写固定电话号码</asp:RegularExpressionValidator></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px; HEIGHT: 21px"></TD>
																						<TD style="WIDTH: 79px; HEIGHT: 21px">
																							<asp:Label id="Label16" runat="server" Font-Size="X-Small">家庭人口</asp:Label></TD>
																						<TD style="WIDTH: 251px; HEIGHT: 21px">
																							<asp:RadioButtonList id="RadioButtonList8" runat="server" Width="216px" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="1">1</asp:ListItem>
																								<asp:ListItem Value="2">2</asp:ListItem>
																								<asp:ListItem Value="3" Selected="True">3</asp:ListItem>
																								<asp:ListItem Value="4">4</asp:ListItem>
																								<asp:ListItem Value="更多">更多</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px; HEIGHT: 23px">
																							<asp:RequiredFieldValidator id="RequiredFieldValidator5" runat="server" ErrorMessage="RequiredFieldValidator"
																								ControlToValidate="TextBox4">*</asp:RequiredFieldValidator></TD>
																						<TD style="WIDTH: 79px; HEIGHT: 23px">
																							<asp:Label id="Label21" runat="server" Font-Size="X-Small">验证码</asp:Label></TD>
																						<TD style="WIDTH: 251px; HEIGHT: 23px">
																							<asp:TextBox id="TextBox4" runat="server" Width="112px" Height="25px"></asp:TextBox>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG style="WIDTH: 40px; HEIGHT: 20px" height="20" alt="" src="draws.aspx" width="40"></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 17px"></TD>
																						<TD style="WIDTH: 79px"></TD>
																						<TD style="WIDTH: 251px"><FONT face="宋体"></FONT></TD>
																						<TD>
																							<asp:Button id="Button3" runat="server" Width="80px" Height="26px" Font-Size="X-Small" Text="下一步>>"
																								BorderStyle="Outset"></asp:Button></TD>
																					</TR>
																				</TABLE>
																			</asp:panel><asp:panel id="Panel2" runat="server" Height="484px" Width="552px" Visible="False">
																				<TABLE id="Table2" style="WIDTH: 584px; HEIGHT: 610px" cellSpacing="1" cellPadding="1"
																					width="584" background="./images/bg1.gif" border="1">
																					<TR>
																						<TD style="WIDTH: 120px; HEIGHT: 39px">&nbsp;
																							<asp:Label id="Label1" runat="server" Font-Size="X-Small">外教性别</asp:Label></TD>
																						<TD style="HEIGHT: 39px">
																							<asp:RadioButtonList id="RadioButtonList9" runat="server" Width="264px" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="男" Selected="True">男</asp:ListItem>
																								<asp:ListItem Value="女">女</asp:ListItem>
																								<asp:ListItem Value="无所谓">无所谓</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px">&nbsp;
																							<asp:Label id="Label15" runat="server" Font-Size="X-Small">是否会说中文</asp:Label></TD>
																						<TD>
																							<asp:RadioButtonList id="RadioButtonList10" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="是">是</asp:ListItem>
																								<asp:ListItem Value="否">否</asp:ListItem>
																								<asp:ListItem Value="无所谓" Selected="True">无所谓</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px">&nbsp;
																							<asp:Label id="Label17" runat="server" Font-Size="X-Small">房间结构</asp:Label></TD>
																						<TD>
																							<asp:RadioButtonList id="RadioButtonList12" runat="server" Width="312px" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="一室一厅">一室一厅</asp:ListItem>
																								<asp:ListItem Value="两室一厅" Selected="True">两室一厅</asp:ListItem>
																								<asp:ListItem Value="三室一厅">三室一厅</asp:ListItem>
																								<asp:ListItem Value="其他">其他</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px">&nbsp;
																							<asp:Label id="Label19" runat="server" Font-Size="X-Small">是否能上网</asp:Label></TD>
																						<TD>
																							<asp:RadioButtonList id="RadioButtonList11" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="是" Selected="True">是</asp:ListItem>
																								<asp:ListItem Value="否">否</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px; HEIGHT: 32px">&nbsp;
																							<asp:Label id="Label2" runat="server" Font-Size="X-Small">是否有宠物</asp:Label></TD>
																						<TD style="HEIGHT: 32px">
																							<asp:RadioButtonList id="RadioButtonList13" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal">
																								<asp:ListItem Value="是" Selected="True">是</asp:ListItem>
																								<asp:ListItem Value="否">否</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px; HEIGHT: 41px">&nbsp;
																							<asp:Label id="Label3" runat="server" Font-Size="X-Small">如何知道我们</asp:Label></TD>
																						<TD style="HEIGHT: 41px">
																							<asp:RadioButtonList id="RadioButtonList14" runat="server" Width="312px" Height="24px" Font-Size="X-Small"
																								RepeatDirection="Horizontal">
																								<asp:ListItem Value="报纸杂志" Selected="True">报纸杂志</asp:ListItem>
																								<asp:ListItem Value="网站">网站</asp:ListItem>
																								<asp:ListItem Value="朋友介绍">朋友介绍</asp:ListItem>
																								<asp:ListItem Value="其他">其他</asp:ListItem>
																							</asp:RadioButtonList></TD>
																					</TR>
																					<TR>
																						<TD style="WIDTH: 120px; HEIGHT: 39px" align="center" colSpan="2">
																							<DIV style="DISPLAY: inline; WIDTH: 152px; HEIGHT: 18px" ms_positioning="FlowLayout">我的住家网会员协议</DIV>
																						</TD>
																					</TR>
																					<TR>
																						<TD style="FONT-SIZE: x-small; LINE-HEIGHT: 25px" colSpan="2">
																							<P>&nbsp;&nbsp;&nbsp;&nbsp;1. 本人保证在本网站提供的资料真实、准确，我的住家网有权进行一切必要的核实。<BR>
																								&nbsp;&nbsp;&nbsp; 2. 本人同意我的住家网将“我的资料（包括图片等）”放在本网站和机构的其它相关网站<BR>
																								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 发布,并由我的住家网和各地的代理机构在海外进行推广。<BR>
																								&nbsp;&nbsp;&nbsp; 3. 我的住家网有权对您的资料进行编辑删减，有权决定接受、保留或取消您的接待资格。<BR>
																								&nbsp;&nbsp;&nbsp; 4. 我的住家网对本网站内容有最终解释权。<BR>
																								&nbsp;以上信息保证真实，如果不实，出现任何后果，全部由填表人负责。<BR>
																								&nbsp;&nbsp;&nbsp;&nbsp; 服务条款提交本页注册信息，即表示您同意我们的服务条款，并且已经阅读并理解我们的隐私&nbsp; 条款。 
																								未经您的同意，我的住家网绝对不会泄漏您的个人隐私信息。</P>
																						</TD>
																					</TR>
																					<TR>
																						<TD colSpan="3">
																							<P><FONT face="宋体">&nbsp;
																									<asp:CheckBox id="CheckBox2" runat="server" Font-Size="X-Small" Text="同意并接受我的住家网会员协议" Checked="True"
																										AutoPostBack="True"></asp:CheckBox>&nbsp; </FONT>
																							</P>
																							<P><FONT face="宋体">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																									<asp:Button id="Button4" runat="server" Width="80px" Height="26px" Font-Size="X-Small" Text="上一步>>"
																										BorderStyle="Outset" CausesValidation="False"></asp:Button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
																									<asp:Button id="Button1" runat="server" Width="80px" Height="26px" Font-Size="X-Small" Text="提交"
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
																		<td colspan="2" class="int1"><p><FONT face="宋体"></FONT>
																			</p>
																		</td>
																	</tr>
																	<tr>
																		<td><FONT face="宋体"></FONT>
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
										<div align="center" style="FONT-SIZE: x-small">&nbsp;电话：&nbsp;010-85493388 
											&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;电子邮箱：<a href="mailto:service@myhomestay.com.cn">service@myhomestay.com.cn</a>
										</div>
									</div>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>
									<div align="center" class="STYLE1">
										<div align="center" style="FONT-SIZE: x-small">&nbsp;客服QQ： <a target="blank" style="TEXT-DECORATION: none" href="http://wpa.qq.com/msgrd?V=1&amp;Uin=401065033&amp;Site=我的住家网0081&amp;Menu=yes">
												<img src="http://wpa.qq.com/pa?p=1:401065033:1" alt="我的住家网0081" border="0" align="absMiddle"
													style="WIDTH: 74px; HEIGHT: 23px" height="23" width="74"> </a>401065033 <a style="TEXT-DECORATION: none" target="blank" href="http://wpa.qq.com/msgrd?V=1&amp;Uin=437488282&amp;Site=我的住家网0082&amp;Menu=yes">
												<img src="http://wpa.qq.com/pa?p=1:437488282:1" alt="我的住家网0082" border="0" align="absMiddle">
											</a>437488282 <a style="TEXT-DECORATION: none" target="blank" href="http://wpa.qq.com/msgrd?V=1&amp;Uin=595574668&amp;Site=我的住家网0083&amp;Menu=yes">
												<img src="http://wpa.qq.com/pa?p=1:595574668:1" alt="我的住家网0083" border="0" align="absMiddle">
											</a>595574668
										</div>
									</div>
								</td>
							</tr>
							<tr>
								<td width="37">&nbsp;</td>
								<td width="883">
									<div align="center" class="STYLE1">
										<div align="center" style="FONT-SIZE: x-small"><font color="#006499" style="FONT-SIZE: x-small">&nbsp;客服MSN：&nbsp;</font>
											<img src="images/MSN_sub.gif" width="19" height="18"><a target="blank" href="msnim:chat?contact=china200701@hotmail.com" class="index">china200701@hotmail.com</a>
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
