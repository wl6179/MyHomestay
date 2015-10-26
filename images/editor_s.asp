<%
dim bdir
bdir=trim(request("blogurl"))
%>
//if (oblog_bLoad==false)
//{
	//oblog_InitDocument("Body","GB2312");
//}
var oblog_bIsIE5=document.all;
var canusehtml='1';
var PostType=1;

if (oblog_bIsIE5){
	var IframeID=frames["oblog_Composition"];
}
else{
	var IframeID=document.getElementById("oblog_Composition").contentWindow;
	var oblog_bIsNC=true;
}

if (oblog_bLoad==false)
{
	oblog_InitDocument("Body","GB2312");
}
function submits(){
document.all("edit").value=IframeID.document.body.innerHTML;
}

function del_space(s)
{
	for(i=0;i<s.length;++i)
	{
	 if(s.charAt(i)!=" ")
		break;
	}

	for(j=s.length-1;j>=0;--j)
	{
	 if(s.charAt(j)!=" ")
		break;
	}

	return s.substring(i,++j);
}

function Verifycomment()
{
	 v = del_space(document.commentform.UserName.value);
     if (v.length == 0)
     {
        alert("您忘了留下名字!");
	return false;
     }
	 	v = del_space(document.commentform.commenttopic.value);
     if (v.length == 0)
     {
        alert("您忘了填写题目!");
	return false;
     }
	 
	submits(); 
	if (document.commentform.edit.value == "")
     {
        alert("内容不能为空!");
	return false;
     }
	 
	var codeid=document.getElementById("CodeStr");
	if(codeid != null){ 
		if (document.commentform.CodeStr.value == "") {
		 alert("验证码不能为空!");
		return false;
		}
	}
	return true;
}

function reply_quote(id)
{
	IframeID.focus();
	oblog_InsertSymbol("<div class='quote'><strong>以下引用"+document.all["n_"+id].innerHTML+"在"+document.all["t_"+id].innerHTML+"发布的评论:</strong><br /><br />"+document.all["c_"+id].innerHTML+"</div><br />")
	IframeID.focus();
}
function DecodeCookie(str)
{
    var strArr;
    var strRtn="";
    strArr=str.split("a");
    try{
        for (var i=strArr.length-1;i>=0;i--)
        strRtn+=String.fromCharCode(eval(strArr[i]));
    }catch(e){
    }
    return strRtn; 
}

if (oblog_bIsNC){
document.write('<iframe width="260" height="165" id="colourPalette" src="<%=bdir%>images/nc_selcolor.htm" style="visibility:hidden; position: absolute; left: 0px; top: 0px;" frameborder="0" scrolling="no" ></iframe>');
}

if (PostType == 0)
{
	oblog_setMode(3);
	document.getElementById("oblog_TabDesign").style.display='none';
	document.getElementById("oblog_TabDesign").style.display='none';
	//onpaste
	//document.selection.createRange().text
	//(window.clipboardData.getData("Text")
}

//数据传递
function oblog_CopyData(hiddenid)
{
	//document.Dvform.Submit.disabled=true;
	//document.Dvform.Submit2.disabled=true;
	if (PostType == 0 && oblog_bTextMode == 3)
	{
		oblog_PasteData()
	}
	d = IframeID.document;
	if (oblog_bTextMode == 2)
	{
		cont = d.body.innerText;
	}else{
		cont = d.body.innerHTML;  
	}
	var ChekEmptyCode = oblog_ChekEmptyCode(cont);
	if (ChekEmptyCode == '' || ChekEmptyCode == null)
	{
		cont='';
	}
	else{
		cont = oblog_correctUrl(cont);
		if (oblog_filterScript)
		cont=oblog_FilterScript(cont);
	}
	document.getElementById(hiddenid).value = cont;
}

function oblog_PasteData()
{
	var regExp;
	cont = IframeID.document.body.innerHTML;
	regExp = /<[s|t][a-z]([^>]*)>/ig
	cont = cont.replace(regExp, '');
	regExp = /<\/[s|t][a-z]([^>]*)>/ig
	cont = cont.replace(regExp, '');
	IframeID.document.body.innerHTML = cont
}
//-------------------------------------
function ctlent(eventobject)
{
	if(event.ctrlKey && event.keyCode==13)
	{
		this.document.Dvform.submit();
	}
}

function putEmot(thenNo)
{
	var ToAdd = '['+thenNo+']';
	IframeID.document.body.innerHTML+=ToAdd;
	IframeID.focus();
}
function gopreview()
{
document.preview.Dvtitle.value=document.Dvform.topic.value;
document.preview.theBody.value=IframeID.document.body.innerHTML;
var popupWin = window.open('', 'preview_page', 'scrollbars=yes,width=750,height=450');
document.preview.submit()
}

//--------------------------------------------------------------------------------

function oblog_foreColor()
{
	if (!oblog_validateMode()) return;
	if (oblog_bIsIE5){
		var arr = showModalDialog("<%=bdir%>images/selcolor.html", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
		if (arr != null) FormatText('forecolor', arr);
		else IframeID.focus();
	}else
		{
		FormatText('forecolor', '');
		//var arr = openEditScript('images/nc_selcolor.htm',250,100)}
		}
}

function oblog_backColor()
{
	if (!oblog_validateMode()) return;
	if (oblog_bIsIE5)
	{
		var arr = showModalDialog("<%=bdir%>images/selcolor.html", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
		if (arr != null) FormatText('backcolor', arr);
		else IframeID.focus();
	}else
		{
		FormatText('backcolor', '');
		}
}

function oblog_correctUrl(cont)
{
	var regExp;
	var url=location.href.substring(0,location.href.lastIndexOf("/")+1);
	cont=oblog_rCode(cont,location.href+"#","#");
	cont=oblog_rCode(cont,url,"");
	cont=oblog_rCode(cont,"<a>　</a>","");
	//regExp = /<a.*href=\"(.*)\"[^>]*>/gi;
	//将连接加上blank标记
	//regExp = /<(a[^>]*) href=([^ |>]*)([^>]*)/gi
	//cont = cont.replace(regExp, "<$1 href=$2 target=\"_blank\" ") ;
	//regExp = /<([^>]*)/gi //转换为小写htm
	//cont = cont.replace(regExp, function($1){return $1.toLowerCase()})
	return cont;
}


function oblog_cleanHtml()
{
	if (oblog_bIsIE5){
	var fonts = IframeID.document.body.all.tags("FONT");
	}else{
	var fonts = IframeID.document.getElementsByTagName("FONT");
	}
	var curr;
	for (var i = fonts.length - 1; i >= 0; i--) {
		curr = fonts[i];
		if (curr.style.backgroundColor == "#ffffff") curr.outerHTML = curr.innerHTML;
	}
}

function oblog_getPureHtml()
{
	var str = "";
	//var paras = IframeID.document.body.all.tags("P");
	//var paras = IframeID.document.getElementsByTagName("p");
	//if (paras.length > 0){
	  //for	(var i=paras.length-1; i >= 0; i--) str= paras[i].innerHTML + "\n" + str;
	//} else {
	str = IframeID.document.body.innerHTML;
	//}
	str=oblog_correctUrl(str);
	return str;
}

function FormatUrl(html)
{
	var regExp = /<a.*href=\"(.*)\"[^>]*>/gi;
	html = html.replace(regExp,"<a href=$1 target=\"_blank\" >")
  return html;
}


function oblog_getEl(sTag,start)
{
	while ((start!=null) && (start.tagName!=sTag)) start = start.parentElement;
	return start;
}

//选择内容替换文本
function oblog_InsertSymbol(str1)
{
	IframeID.focus();
	if (oblog_bIsIE5) oblog_selectRange();
	oblog_edit.pasteHTML(str1);
}


//选择事件
function oblog_selectRange(){
	oblog_selection =	IframeID.document.selection;
	oblog_edit		=	oblog_selection.createRange();
	oblog_RangeType =	oblog_selection.type;
}

//应用html
function oblog_specialtype(Mark1, Mark2){
	var strHTML;
	if (oblog_bIsIE5){
		oblog_selectRange();
		if (oblog_RangeType == "Text"){
			if (Mark2==null)
			{
				strHTML = "<" + Mark1 + ">" + oblog_edit.htmlText + "</" + Mark1 + ">"; 
			}else{
				strHTML = Mark1 + oblog_edit.htmlText +  Mark2; 
			}
			oblog_edit.pasteHTML(strHTML);
			IframeID.focus();
			oblog_edit.select();
		}
		else{window.alert("请选择相应内容！")}	
	}
	else{
		if (Mark2==null)
		{
		strHTML	=	"<" + Mark1 + ">" + IframeID.document.body.innerHTML + "</" + Mark1 + ">"; 
		}else{
		strHTML = Mark1 + IframeID.document.body.innerHTML +  Mark2; 
		}
		IframeID.document.body.innerHTML=strHTML
		IframeID.focus();
	}
}

// 修改编辑栏高度
function oblog_Size(num)
{
	var obj=document.getElementById("oblog_edit");
	//if (parseInt(obj.style.height)+num>=300) {
		//alert(obj.style.height)
		//obj.style.height = (parseInt(obj.style.height) + num);
	if (num>0){
		obj.style.height=num+"px";
		obj.style.width="100%";
		}
	else{
		obj.style.height="";
		//alert(-num+"px");
		obj.style.width=-num+"px";
	}
}

function oblog_getText()
{
	if (oblog_bTextMode==2)
		return IframeID.document.body.innerText;
	else
	{
		oblog_cleanHtml();
		return IframeID.document.body.innerHTML;
	}
}

function oblog_putText(v)
{
	if (oblog_bTextMode==2)
		IframeID.document.body.innerText = v;
	else
		IframeID.document.body.innerHTML = v;
}
function oblog_doSelectClick(str, el)
{
	var Index = el.selectedIndex;
	if (Index != 0){
		el.selectedIndex = 0;
		FormatText(str,el.options[Index].value);
	}
}
//查找配对字符出现次数,没有结果为0
function TabCheck(word,str){
	var tp=0
	chktp=str.search(word);
	if (chktp!=-1)
	{
	eval("var tp=\""+str+"\".match("+word+").length")
	}
	return tp;
}

//Colour pallete top offset
function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}

//Colour pallete left offset
function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}

//Function to hide colour pallete
function hideColourPallete() {
	document.getElementById("colourPalette").style.visibility="hidden";
}


//------------------------------------------------------
function OpenSmiley()
{
	var arr = showModalDialog("<%=bdir%>images/smiley.htm", "", "dialogWidth:60em; dialogHeight:15.5em; status:0; help:0");
	if (arr != null){
		var ss;
		ss=arr.split("*")
		path=ss[0];
		ubbstring=ss[1];
		IframeID.document.body.innerHTML+=ubbstring;
	}
	else IframeID.focus();
}



function rand() {
	return parseInt((1000)*Math.random()+1);
}


//图片与链接事件
function oblog_UserDialog(what)
{
	if (!oblog_validateMode()) return;
	IframeID.focus();
	if (what == "CreateLink") {
		if (oblog_bIsNC)
		{
			insertLink = prompt("请填写超级链接地址信息：", "http://");			
			if ((insertLink != null) && (insertLink != "") && (insertLink != "undefined")) {
			IframeID.document.execCommand('CreateLink', false, insertLink);
			}else{
			IframeID.document.execCommand('unlink', false, null);
			}
		}
		else {
			IframeID.document.execCommand(what, true, null);
		}
	}
	//去掉添加图片时的src="file://
	if(what == "InsertImage"){
		imagePath = prompt('请填写图片链接地址信息：', 'http://');			
		if ((imagePath != null) && (imagePath != "")) {
			IframeID.document.execCommand('InsertImage', false, imagePath);
		}
		IframeID.document.body.innerHTML = (IframeID.document.body.innerHTML).replace("src=\"file://","src=\"");
	}
	oblog_pureText = false;
	IframeID.focus();
}

//--------------------
function oblog_GetRangeReference(editor)
{
	editor.focus();
	var objReference = null;
	var RangeType = editor.document.selection.type;
	var selectedRange = editor.document.selection.createRange();
	
	switch(RangeType)
	{
	case 'Control' :
		if (selectedRange.length > 0 ) 
		{
			objReference = selectedRange.item(0);
		}
	break;
	case 'None' :
		objReference = selectedRange.parentElement();
		break;
	case 'Text' :
		objReference = selectedRange.parentElement();
		break;
	}
	return objReference
}

function oblog_CheckTag(item,tagName)
{
	if (item.tagName.search(tagName)!= -1)
	{
		return item;
	}
	if (item.tagName == 'BODY')
	{
		return false;
	}
	item=item.parentElement;
	return oblog_CheckTag(item,tagName);
}

function oblog_code()
{
	oblog_specialtype("<div class=HtmlCode style='cursor: pointer'; title='点击运行该代码！' onclick=\"preWin=window.open('','','');preWin.document.open();preWin.document.write(this.innerText);preWin.document.close();\">","</div>");	
	//oblog_specialtype("<div class=HtmlCode>","</div>");	
}

function oblog_quote()
{
	oblog_specialtype("<div style='margin:5px 20px;border:1px solid #CCCCCC;padding:5px; background:#F3F3F3'>","</div>");
}

function oblog_replace()
{
	var arr = showModalDialog("<%=bdir%>images/replace.html", "", "dialogWidth:16.5em; dialogHeight:13em; status:0; help:0");
	if (arr != null){
		var ss;
		ss = arr.split("*")
		a = ss[0];
		b = ss[1];
		i = ss[2];
		con = IframeID.document.body.innerHTML;
		if (i == 1)
		{
			con = oblog_rCode(con,a,b,true);
		}else{
			con = oblog_rCode(con,a,b);
		}
		IframeID.document.body.innerHTML = con;
	}
	else IframeID.focus();
}

function insertSpecialChar()
{
	var arr = showModalDialog("<%=bdir%>images/specialchar.html", "","dialogWidth:25em; dialogHeight:15em; status:0; help:0");
	if (arr != null) oblog_InsertSymbol(arr);
	IframeID.focus() ;
}

function doZoom( sizeCombo ) 
{
	if (sizeCombo.value != null || sizeCombo.value != "")
	if (oblog_bIsIE5){
	var z = IframeID.document.body.runtimeStyle;}
	else{
	var z = IframeID.document.body.style;
	}
	z.zoom = sizeCombo.value + "%" ;
}
//--------------------


function oblog_InsertSymbol(str1)
{
	oblog_Composition.focus();
	if (oblog_bIsIE5) oblog_selectRange();	
	oblog_edit.pasteHTML(str1);
}

function oblog_foremot()
{
	var arr = showModalDialog("<%=bdir%>images/emot.htm", "", "dialogWidth:26em; dialogHeight:13em; status:0; help:0");
	
	if (arr != null)
	{
		//content=oblog_Composition.document.body.innerHTML;
		//content=content+arr;
		//oblog_Composition.document.body.innerHTML=content;
		oblog_InsertSymbol(arr);
		IframeID.focus();
	}
	else IframeID.focus();
}

function oblog_forlink()
{
if (oblog_bIsIE5){		
	var arr=showModalDialog("<%=bdir%>images/link.htm",window, "dialogWidth:23em; dialogHeight:11em; status:0; help:0");
	IframeID.focus();
	if (arr != null)
	{		
		oblog_InsertSymbol(arr);
		IframeID.focus();
		
	}
	else IframeID.focus();
}
else {oblog_UserDialog('CreateLink');}
}

function getHTML() {
	var html;
	if (!oblog_bTextMode) 
	{
	html = IframeID.document.body.innerHTML
	}
	else
	{
	html = IframeID.document.body.innerText
	}
	return html;
}