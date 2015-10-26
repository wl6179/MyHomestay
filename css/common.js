
function writeFlash(width,height,url)
{
	document.write("<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0'  width = "+width+" height = "+height+">");
	document.write("<PARAM NAME=movie VALUE='"+url+"'>"); 
	document.write("<PARAM NAME=quality VALUE=high>");
	document.write("<PARAM NAME=bgcolor VALUE=#FFFFFF>");
	document.write("<PARAM NAME=wmode VALUE=transparent>");
	document.write("<EMBED src='"+url+"' quality=high bgcolor=#FFFFFF wmode=transparent width = "+width+" height = "+height+"  TYPE='application/x-shockwave-flash' PLUGINSPAGE='http://www.macromedia.com/go/getflashplayer'></EMBED>");
	document.write("</OBJECT>");
}

function writeMovie(width,height,url)
{
	document.write("<object id='TVplayer' classid='clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95' codebase='http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701' standby='Loading Microsoft Windows Media Player components...' width='"+width+"' height='"+height+"' type='application/x-oleobject ' align='center'>");
	document.write("<param name='animationatstart' value='0'>");
	document.write("<param NAME='autostart' value='1'>");
	document.write("<param NAME='AutoSize' value='1'>");
	document.write("<param NAME='EnableContextMenu' Value='0'>");
	document.write("<param NAME='playCount' value='-1'>");
	document.write("<param NAME='ShowStatusBar' VALUE='1'>");
	document.write("<param NAME='ShowControls' VALUE='0'>");
	document.write("<param NAME='ClickToPlay' VALUE='0'>");
	document.write("<param NAME='TransparentAtStart' VALUE='0'>");
	document.write("<param name='DisplaySize' value='0'>");
	document.write("<param name='Filename' value='"+url+"");
	document.write("</object>");
}

function writeHtml(v_html){
	document.write(v_html);
}

function goUrl(name, mode) {
	if (eval(name) == '') {
		alert("");
		return;	
	} else {
		if (!mode) {
    		document.location.href = eval(name);
		}
	}	
}	

// Menu
function GoMenu(name, mode) {
  if (eval(name) == '') {
		alert("Content comming soon.");
		return;
	} else {
    if (!mode) {
    		document.location.href = eval(name);
    } else {
			var values = mode.split('|');
			if(values.length == 2) {
				window.open(eval(name), '', 'width='+eval(values[0])+',height='+eval(values[1]));
			} else if(values.length > 2) {
				window.open(eval(name), '', 'width='+eval(values[0])+',height='+eval(values[1])+","+values[2]);
			} else {
    		window.open(eval(name), '');
			}
  	}
	}
}
function mother() {
document.writeln("<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'")	
document.writeln(" codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0'")	
document.writeln(" WIDTH='234' HEIGHT='428' id='#' ALIGN=''>")	
document.writeln(" <PARAM NAME=movie VALUE='images/flash/index_left.swf'> ")	
document.writeln(" <PARAM NAME=quality VALUE=high> ")	
document.writeln(" <PARAM NAME=bgcolor VALUE=#FFFFFF>")	
document.writeln(" <PARAM NAME=wmode VALUE=transparent>")	
document.writeln("</OBJECT>")	
}

function Top(){
document.writeln("<a name='top'></a>")	
document.writeln("<script language='JavaScript'>Top_menu();</script>")	
document.writeln("<table border='0' cellpadding='0' cellspacing='0' width='100%'>")
document.writeln("<tr>")
document.writeln("	<td valign='bottom' width='44' background='"+str_link+"/img/img_left.jpg'><img src='"+str_link+"/img/img_left_b.gif' width='44' height='22' alt='' border='0'></td>")
document.writeln("	<td valign='top' bgcolor='#FFFFFF' width='100%'>")
document.writeln("		<table border='0' cellpadding='0' cellspacing='0' width='100%'>")
document.writeln("		<tr>")
document.writeln("			<td class='bg' colspan='2'><img src='"+str_link+"/img/img_top.jpg' width='934' height='45' alt='' border='0'></td>")
document.writeln("		</tr>")
document.writeln("		<tr>")
document.writeln("			<td width='190' valign='top'>")
document.writeln("				<table border='0' cellpadding='0' cellspacing='0' width='190'>")
document.writeln("				<tr>")
document.writeln("					<td><img src='"+str_link+"/img/round_t.gif' width='20' height='20' alt='' border='0'></td>")
document.writeln("				</tr>")
document.writeln("				<tr>")
document.writeln("					<td height='66' align='center'><a href=javascript:goUrl('home')><img src='"+str_link+"/img/logo.gif' width='145' height='25' alt='' border='0'></a></td>")
document.writeln("				</tr>")
document.writeln("				</table>")
document.writeln("			</td>")
document.writeln("			<td height='93'></td>")
document.writeln("		</tr>")
document.writeln("		</table>")
document.writeln("		<table border='0' cellpadding='0' cellspacing='0' width='856' height='400'>")
document.writeln("		<tr>")
document.writeln("			<td width='190' valign='top'>")
document.writeln("				<!--Left S.-->")				
document.writeln("				<script language='JavaScript'>Left_menu();</script>")		
document.writeln("				<!--Left E.-->")	
document.writeln("			</td>")
document.writeln("			<td width='666'  valign='top'>")
document.writeln("			<!-- Body S.-->")
}
function Top2(){
document.writeln("<a name='top'></a>")	
document.writeln("<script language='JavaScript'>Top_menu();</script>")	
document.writeln("<table border='0' cellpadding='0' cellspacing='0' width='100%'>")
document.writeln("<tr>")
document.writeln("	<td valign='bottom' width='44' background='"+str_link+"/img/img_left.jpg'><img src='"+str_link+"/img/img_left_b.gif' width='44' height='22' alt='' border='0'></td>")
document.writeln("	<td valign='top' bgcolor='#FFFFFF' width='100%'>")
document.writeln("		<table border='0' cellpadding='0' cellspacing='0' width='100%'>")
document.writeln("		<tr>")
document.writeln("			<td class='bg' colspan='2'><img src='"+str_link+"/img/img_top.jpg' width='934' height='45' alt='' border='0'></td>")
document.writeln("		</tr>")
document.writeln("		<tr>")
document.writeln("			<td width='190' valign='top'>")
document.writeln("				<table border='0' cellpadding='0' cellspacing='0' width='190'>")
document.writeln("				<tr>")
document.writeln("					<td><img src='"+str_link+"/img/round_t.gif' width='20' height='20' alt='' border='0'></td>")
document.writeln("				</tr>")
document.writeln("				<tr>")
document.writeln("					<td height='66' align='center'><a href=javascript:goUrl('home')><img src='"+str_link+"/img/logo.gif' width='145' height='25' alt='' border='0'></a></td>")
document.writeln("				</tr>")
document.writeln("				</table>")
document.writeln("			</td>")
document.writeln("			<td height='93'></td>")
document.writeln("		</tr>")
document.writeln("		</table>")
document.writeln("		<table border='0' cellpadding='0' cellspacing='0' width='856' height='400'>")
document.writeln("		<tr>")
document.writeln("			<td width='190' valign='top'>")
document.writeln("				<!--Left S.-->")				
document.writeln("				<script language='JavaScript'>Left_menu2();</script>")		
document.writeln("				<!--Left E.-->")	
document.writeln("			</td>")
document.writeln("			<td width='666'  valign='top'>")
document.writeln("			<!-- Body S.-->")
}
function space(){
document.writeln("<tr><td height='60'></td></tr>")
}

function section_icon_pr(){
document.writeln("<table border='0' cellpadding='0' cellspacing='0' width='100%' >")
document.writeln("<tr>")
document.writeln("<td><img src='"+str_link+"/img/blank.gif' alt='' border='0' width='1' height='59'></td>")
document.writeln("</tr>")
document.writeln("</table>")
}

function section_icon_about(){
document.writeln("<table border='0' cellpadding='0' cellspacing='0' width='100%' >")
document.writeln("<tr>")
document.writeln("<td><img src='"+str_link+"/img/blank.gif' alt='' border='0' width='1' height='59'></td>")
document.writeln("</tr>")
document.writeln("</table>")
}

function section_icon_customer(){
document.writeln("<table border='0' cellpadding='0' cellspacing='0' width='100%' >")
document.writeln("<tr>")
document.writeln("<td><img src='"+str_link+"/img/blank.gif' alt='' border='0' width='1' height='59'></td>")
document.writeln("</tr>")
document.writeln("</table>")
}

function section_icon_open(){
document.writeln("<table border='0' cellpadding='0' cellspacing='0' width='100%' >")
document.writeln("<tr>")
document.writeln("<td><img src='"+str_link+"/img/blank.gif' alt='' border='0' width='1' height='59'></td>")
document.writeln("</tr>")
document.writeln("</table>")
}

function Footer(){
document.writeln("<DIV style='position:relative; visibility:show; left:0px; top:-59px; z-index:1;'>")
document.writeln("<a href='#top'><img src='"+str_link+"/img/top.gif' width='14' height='11' alt='' border='0' style='position:absolute; visibility:show; left:900px; top:0px;'></a><img src='"+str_link+"/img/blank.gif' width='10' height='48' alt='' border='0' style='position:absolute; visibility:show; left:892px; top:11px;'>")
document.writeln("</DIV>")
document.writeln("<table border='0' cellpadding='0' cellspacing='0' width='100%'>")
document.writeln("<tr>")
document.writeln("	<td width='44'><img src='"+str_link+"/img/spacer.gif' width='41' height='1' alt='' border='0'></td>")
document.writeln("	<td width='100%' background='"+str_link+"/img/bg_footer.gif'><table border='0' cellpadding='0' cellspacing='0' width='860'><tr><td><img src='"+str_link+"/img/copyright_2005_09_15.gif' width='538' height='48' alt='' border='0' usemap='#map_copy'><map name='map_copy'><area shape='rect' coords='469,25,537,38' href='"+str_link+"/sitemap/private.html'></map></td><td align='right' valign='top'></td></tr></table></td>")
document.writeln("</tr>")
document.writeln("</table>")
}



//Äü¸Þ´º
var mover = null
var object = "qmenu"
var v_scrollHeight = "";


function check() {
		if(""==v_scrollHeight) {
			v_scrollHeight = document.body.scrollHeight;
		}
//		document.test.test2.value=(v_scrollHeight < (document.body.scrollTop+375+190))+","+document.body.scrollHeight+","+v_scrollHeight+","+(v_scrollHeight-375+190)+","+document.body.height;
		if (navigator.appName.indexOf("Netscape") != -1) {
				var mover = "document.qmenu.top = window.pageYOffset"
		}
		else {
				var mover = "qmenu.style.pixelTop = document.body.scrollTop+198";
				if(v_scrollHeight < (document.body.scrollTop+375+190)) {
					mover = "qmenu.style.pixelTop = "+(v_scrollHeight-(375));
				}else {
					mover = "qmenu.style.pixelTop = document.body.scrollTop+198";
				}

		}

        object = "qmenu"
        eval(mover)
        setTimeout("check()", 100)
}
function Quick_menu(){	
document.writeln("	<!-- Body E.-->	")
document.writeln("			</td>")
document.writeln("			<td width='1' background='"+str_link+"/img/dot_v.gif' valign='top'><img src='"+str_link+"/img/blank.gif' width='1' height='40' alt='' border='0'>")
document.writeln("		</tr>")
document.writeln("		</table>")
document.writeln("	</td>")
document.writeln("</tr>")
document.writeln("</table>")
document.writeln("<DIV id='qmenu' style='position:absolute; visibility:show; left:896px; top:198px; z-index:2'>")	
document.writeln("<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'	 codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0' WIDTH='94' HEIGHT='375' id='#' ALIGN=''>")
document.writeln("<PARAM NAME=movie value='"+str_link+"/img/flash/quick.swf'>")
document.writeln("<PARAM NAME=quality VALUE=high>")
document.writeln("<PARAM NAME=bgcolor VALUE=#FFFFFF>")
document.writeln("<PARAM NAME=wmode VALUE=transparent> ")
document.writeln("</OBJECT>")
document.writeln("</DIV>")
}



//Popup 		
function popup(Fn,Win, X, Y, Scroll){
l = (screen.width) ?	(screen.width- X) / 2	: 0;
t = (screen.height) ?	(screen.height- Y) / 2 : 0;	
	NewWindow=window.open(Fn,Win,'width='+X+',height='+Y+',top='+t+',left='+l+',scrollbars='+Scroll+',toolbar=no,location=no,directories=no,status=no,resizable=no,menubar=no');
}

function winopen(theUrl,winSize){
  window.open(theUrl,"",winSize + "status=no, resizable=no, scrollbars=no");
}

function sc_winopen(theUrl,winSize){
  window.open(theUrl,"",winSize + ",height=500, status=no, resizable=no, scrollbars=yes");
}
function sc_winopen2(theUrl,winSize){
  window.open(theUrl,"", "width=660,height=500, status=no, resizable=no, scrollbars=yes");
}

function Open_TV(theURL) {
  var height = screen.height; 
  var width = screen.width; 
  var leftpos = width / 2 - 300; 
  var toppos = height / 2 - 225;
  pop_info = window.open(theURL,'TV','width=390,height=550,left='+leftpos+',top='+toppos+'');
  pop_info.focus();
}

function move_site(frm) {
    var myindex=frm.selectedIndex
    if (frm.options[myindex].value != null) {
				 document.location.href(frm.options[myindex].value);
    }
}

function Open_Webzine(theURL) {

	if(parseInt(theURL) < 195){
		OpenEBook.OpenBook(theURL,'ebook');
	}
	else
		if (parseInt(theURL) < 200000){
		OpenEBook.OpenBook(theURL);
	}else{
	  pop_info = window.open('http://davadb1.kepco.co.kr/news/webzine/'+theURL+'/webzine.html','webzine','navigation toolbar=yes,location=yes,status bar=yes,scrollbars=yes,resizable=yes,width=800,height=600,left=0,top=0');
	  pop_info.focus();
	}
}


function monthly(){
  OpenEBook.OpenBook(146);
}
function culture(){
	Open_Select('../down/culture/2006_0506/webbook.html');
}

function magazine(){
	win_life();
}

function safe(){

	Open_Cyber('http://www.kepco.co.kr/ele_safe/cyber_ele/2006_05/init.htm');
}

function open_brief(){
	 window.open('http://news.go.kr/warp/webapp/news/list?category_id=p_mini_news', '', 'scrollbars=yes,resizable=yes,toolbar=yes,menubar=yes,location=no,top=0,left=0');
	if ( notice_getCookie( "pop_0520" ) != "done" ) 
	{
		  window.open('/brief/popup_0520.html', '', 'scrollbars=no,resizable=no,toolbar=no,menubar=no,location=no,width=341,height=490,top=0,left=0');
	}	 
}

function Gomuseum() {
		window.open('/img/flash/museum/main.html' , 'intro' , 'width=760,height=400,left=200,top=200,fullscreen=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0');
	}

function Go765kv() {
		window.open('http://www.kepco.co.kr/765k/' , '765kv' , '');
}

function Gokepcostory() {
		window.open('/img/flash/kepco_story/2006_05/movie.htm' , 'intro' , 'width=640,height=517,left=200,top=100,fullscreen=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0');
}
function Gokepcostory_2006_05() {
		window.open('/img/flash/kepco_story/main.html' , 'intro' , 'width=700,height=400,left=200,top=200,fullscreen=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0');
}

function pop_20060513() {
		read.src="http://cis.kepco.co.kr/cis/pr/notice/notice_read.jsp?sn=75";
		window.open('http://cis.kepco.co.kr/csagent/cyber_spot_new/jubu/jubu_lot_search.html' , 'intro' , 'width=550,height=380,left=200,top=200,fullscreen=0,location=0,menubar=0,scrollbars=0,status=0,toolbar=0');
}



function Open_Movie(theURL) {
  var height = screen.height;
  var width = screen.width; 
  var leftpos = width / 2 - 175;
  var toppos = height / 2 - 200;
  pop_info = window.open('/pr/down/cf/d4_03_pop_'+theURL+'.html','TV','width=390,height=450,left='+leftpos+',top='+toppos+'');
  pop_info.focus();
}


function public_wisdom(){
	 window.open('http://www.kemco.or.kr/e-saving/event/event_01.asp?left_menu_num=07&line_map=01&b_eventcode=event1', 'aaa');
}

function banner_net(){
	 window.open('http://event.media.daum.net/mocie/500.html', 'ddd', 'scrollbars=no,resizable=no,toolbar=no,menubar=no,location=no,width=330,height=335,top=0,left=0');
}

function openMap()
{
	SimpleCenterWindow('http://cis.kepco.co.kr/csagent/cyber_spot_new/map/map_frame.html', '_map', 753, 715, 0);
}

function SimpleCenterWindow(url,name,width,height,top,left)
{
	if (top == null ) top		= (screen.height - height) / 2;
	if (left == null) left	= (screen.width - width) / 2;

	var centerWindow = window.open(url, name, 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width='+width+',height='+height+',top='+top+',left='+left);

centerWindow.focus();
}

function Open_Cyber(theURL) {

  pop_info = window.open(theURL,'Cyber','');
  pop_info.focus();
}
