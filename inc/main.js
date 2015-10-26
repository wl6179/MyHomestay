function bbimg(o){
	var zoom=parseInt(o.style.zoom, 10)||100;zoom+=event.wheelDelta/12;if (zoom>0) o.style.zoom=zoom+'%';
	return false;
}
function GetCookie (name) {
var CookieFound = false;
var start = 0;
var end = 0;
var CookieString = document.cookie;
var i = 0;

while (i <= CookieString.length) {
start = i ;
end = start + name.length;
if (CookieString.substring(start, end) == name){
CookieFound = true;
break; 
}
i++;
}

if (CookieFound){
start = end + 1;
end = CookieString.indexOf(";",start);
if (end < start)
end = CookieString.length;
return unescape(CookieString.substring(start, end));
}
return "";
}

function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}

function chkdiv(divid){
	var chkid=document.getElementById(divid);
	if(chkid != null){return true; }
	else {return false; }
}

function getpara(){
	var str,pos,parastr
	str = window.location.href;
	pos = str.indexOf("?");
	parastr = str.substring(pos+1);
	return parastr;
}

function oblog_ViewCode(rnum)
{
	var bodyTag="<html><head><style type=text/css>.quote{margin:5px 20px;border:1px solid #CCCCCC;padding:5px; background:#F3F3F3 }\nbody{boder:0px}.HtmlCode{margin:5px 20px;border:1px solid #CCCCCC;padding:5px;background:#FDFDDF;font-size:14px;font-family:Tahoma;font-style : oblique;line-height : normal ;font-weight:bold;}\nbody{boder:0px}</style></head><BODY bgcolor=\"#FFFFFF\" >";
	bodyTag+=document.getElementById('scode'+rnum).value
	bodyTag+="</body></html>"
	preWin=window.open('preview','','left=0,top=0,width=550,height=400,resizable=1,scrollbars=1, status=1, toolbar=1, menubar=0');
	preWin.document.open();
	preWin.document.write(bodyTag);
	preWin.document.close();
	preWin.document.title="查看代码内容";
	preWin.document.charset="UTF-8";
}