/*如果发现自己被嵌套到IFRAME了就替换父窗口的URL,bbs除外;*/
var url=window.location.href;
if(window!=parent&&url.indexOf("bbs")==-1&&url.indexOf("iframe")==-1)
	parent.navigate(url);
function $(x){
    return document.getElementById(x);
}
/*以下显示logo*/
var objLogoBg=$("headLogoBg");
var objLogo=$("headLogo");
if(objLogoBg)
	objLogoBg.innerHTML='<a href="http://www.myhomestay.com.cn/" target="_top"><img src="ShoppingImages/logo.gif" alt="芭芭拉旗下的在线购物网站，快乐、时尚、共享！" /></a>';
else if(objLogo) 
	objLogo.innerHTML='<a href="http://www.myhomestay.com.cn/" target="_top"><img src="ShoppingImages/logoV3A.gif" alt="芭芭拉旗下的在线购物网站，快乐、时尚、共享！" /></a>';

/*
if($("headNote")){
	$("headNote").innerHTML="为了更好地为拍友服务，拍拍网将于2月27日晚 23:30至28日晨9：30进行系统维护，届时拍拍网的交易相关功能将暂时无法使用。<br />由此给您造成不便，我们深表歉意，感谢您对拍拍网的支持！";
	$("headNote").style.display="";
}
*/

function showFlash(showId,flashSrc,flashW,flashH){
	$(showId).innerHTML='<embed src="'+flashSrc+'" height="'+flashH+'" width="'+flashW+'" wmode="transparent" type="application/x-shockwave-flash"></embed>';
}

function setHeadFmCookie(name){
    var cookieValue = "";
    if(document.cookie.length > 0)
    {
        var vec = document.cookie.split(';');
        var search = name + "=";
        for(var i = 0; i < vec.length; ++i)
        {
            vec[i] = vec[i].replace(/(^\s*)|(\s*$)/g, "");
            var offset = vec[i].indexOf(search);
            if(offset != 0)
                continue;
            cookieValue = unescape(vec[i].substring(offset + search.length));
            break;
        }
    }
    return cookieValue;
}

var hs = setHeadFmCookie("hs");
var sNickName= hs.replace(/[0,1]\/\d+\/[0,1]\//,"");
var showStatus="";
var objUser=$("headUser");
var objInto=$("headInto");
var objStatus=$("headStatus");
if(sNickName){
	var sNickName_short = (sNickName.length > 5)?sNickName.substring(0,3)+"...":sNickName;
	showUser="欢迎您，"+sNickName_short+"！";
	if(objUser) objUser.innerHTML=showUser;
	if(objInto){
		if(url.indexOf("my.paipai.com")==-1)
			objInto.innerHTML="[退出]";
		else objInto.innerHTML="退出";
		objInto.href="http://member.paipai.com/cgi-bin/c2cUser_LoginOut";
	}
	var showBasic=hs.split("/")[0];
	var showPayNum=hs.split("/")[1];
	var showTenPay=hs.split("/")[2];
	var objStatus=$("headStatus");
	if(objStatus){
		if(showPayNum>0)
			showStatus="<a href='http://my.paipai.com/cgi-bin/qshop_manage?urltype=1' class='linkBlue'>您有 <b style='color:#ff4e00;text-decoration:underline'>"+showPayNum+"</b>&nbsp;&nbsp;笔交易待付款！</a>";
		else if(showBasic==0)
			showStatus="<a href='http://my.paipai.com/cgi-bin/qshop_manage?urltype=18' class='linkBlue'>您的基本资料尚未设置！</a>";
		else if(showTenPay==0)
			showStatus="<a href='http://pay.paipai.com/cgi-bin/pay_query_user_qpayaccount' target='_blank' class='linkBlue'>您的财付通帐户尚未开通！</a>";
		objStatus.innerHTML=showStatus;
	}
	RememberWelcomeMsg(showStatus);
}

function RememberWelcomeMsg(value){
 document.cookie =  "WelcomeMsg=" + escape(value)+";domain=my.paipai.com";
}

function cleanKeyWord(keyDefault,keyId){
	var obj=$(keyId);
	if(keyDefault==obj.value)
		obj.value="";
}
function checkHeadSubmit(keyDefault,keyId){
	var obj=$(keyId);
	if(keyDefault==obj.value||obj.value=="")
		obj.value="";
	return true;
}

function showHeadOver(obj){
	obj.src=obj.src.replace(".gif","Over.gif");
}

function showHeadOut(obj){
	obj.src=obj.src.replace("Over","");
}

function showMenu(menuName,isIndex){
	var obj=$("menu"+menuName);
	var objList=$("menuList");
	if(obj){
		obj.className="menuHere";
		if(menuName=="Home"){
			obj.style.backgroundImage="url(ShoppingImages/menuHere***.gif)";
			obj.style.paddingLeft="0px";
		}
	}
}

function readCookie(name){
    var cookieValue = "";
    if(document.cookie.length > 0)    {
        var vec = document.cookie.split(';');
        var search = name + "=";
        for(var i = 0; i < vec.length; ++i){
            vec[i] = vec[i].replace(/(^\s*)|(\s*$)/g, "");
            var offset = vec[i].indexOf(search);
            if(offset != 0)
                continue;
            cookieValue = unescape(vec[i].substring(offset + search.length));
            break;
        }
    }
    return cookieValue;
}

function goShopCar(){
	var sUrl = "http://auction.paipai.com/cgi-bin/shopcar?"+Math.random();
	var oPopUp = window.open(sUrl,"ppShopCar");
	oPopUp.focus();	
}

function ShowShopcarNum(){		
	var oShopCarNum =$("spcartnum");
	var oBBSFloat   =$("bbsfloat");
	var iNum = parseInt(readCookie("spcart"));
	if(isNaN(iNum))
		iNum = 0;
	if(oShopCarNum && iNum>0){
		oShopCarNum.innerHTML = "<span class='fontOrange number'>["+iNum+"]</span>";
	}
}

ShowShopcarNum();

var statusNote="拍拍网 - 快乐、时尚、共享";
var isScreen=1;
var isIE6=1;
if(screen.width==1024) isScreen=0;
var nav=navigator.appVersion;
if (nav.indexOf('MSIE')!=-1){
	var curIE=nav.substr(nav.indexOf('MSIE')+5,3);
	if (curIE==6.0) isIE6=0;
}
if(isIE6||isScreen) statusNote+=" - 使用";
if(isIE6) statusNote+="IE6.0";
if(isIE6&&isScreen) statusNote+="以及";
if(isScreen) statusNote+="1024×768分辨率";
if(isIE6||isScreen) statusNote+="访问本站能够获得最佳浏览效果。";
window.status=statusNote;	