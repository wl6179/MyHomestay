var Quote = 0;
var Bold  = 0;
var Italic = 0;
var Underline = 0;
var Code = 0;
var Center = 0;
var Strike = 0;
var Sound = 0;
var Swf = 0;
var Ra = 0;
var Rm = 0;
var Marquee = 0;
var Fly = 0;
var fanzi=0;
var text_enter_url      = "������������ַ";
var text_enter_txt      = "����������˵��";
var text_enter_image    = "������ͼƬ��ַ";
var text_enter_sound    = "�����������ļ���ַ";
var text_enter_swf      = "������FLASH������ַ";
var text_enter_ra      = "������Real������ַ";
var text_enter_rm      = "������RealӰƬ��ַ";
var text_enter_wmv      = "������MediaӰƬ��ַ";
var text_enter_wma      = "������Media������ַ";
var text_enter_mov      = "������QuickTime������ַ";
var text_enter_sw      = "������shockwave������ַ";
var text_enter_email    = "�������ʼ���ַ";
var error_no_url        = "������������ַ";
var error_no_txt        = "����������˵��";
var error_no_title      = "������������ҳ����";
var error_no_email      = "�����������ʼ���ַ";
var error_no_gset       = "������ȷ���ո�ʽ���룡";
var error_no_gtxt       = "�����������֣�";
var text_enter_guang1   = "���ֵĳ��ȡ���ɫ�ͱ߽��С";
var text_enter_guang2   = "Ҫ����Ч�������֣�";
function commentWrite(NewCode) {
document.oblogform.ubbedit.value+=NewCode;
document.oblogform.ubbedit.focus();
return;
}
function storeCaret(text) { 
	if (text.createTextRange) {
		text.caretPos = document.selection.createRange().duplicate();
	}
        if(event.ctrlKey && window.event.keyCode==13){i++;if (i>1) {alert('�������ڷ����������ĵȴ���');return false;}this.document.form.submit();}
}
function AddText(text) {
	if (document.oblogform.ubbedit.createTextRange && document.oblogform.ubbedit.caretPos) {      
		var caretPos = document.oblogform.ubbedit.caretPos;      
		caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
		text + ' ' : text;
	}
	else document.oblogform.ubbedit.value += text;
	document.oblogform.ubbedit.focus(caretPos);
}
function inputs(str)
{
AddText(str);
}
function Curl() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_url, "http://");
var enterTxT   = prompt(text_enter_txt, enterURL);
if (!enterURL)    {
FoundErrors += "\n" + error_no_url;
}
if (!enterTxT)    {
FoundErrors += "\n" + error_no_txt;
}
if (FoundErrors)  {
alert("����"+FoundErrors);
return;
}
var ToAdd = "[URL="+enterURL+"]"+enterTxT+"[/URL]";
document.oblogform.ubbedit.value+=ToAdd;
document.oblogform.ubbedit.focus();
}
function Cimage() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_image, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("����"+FoundErrors);
return;
}
var ToAdd = "[IMG]"+enterURL+"[/IMG]";
document.oblogform.ubbedit.value+=ToAdd;
document.oblogform.ubbedit.focus();
}
function Cemail() {
var emailAddress = prompt(text_enter_email,"");
if (!emailAddress) { alert(error_no_email); return; }
var ToAdd = "[EMAIL]"+emailAddress+"[/EMAIL]";
commentWrite(ToAdd);
}
function Ccode() {
if (Code == 0) {
ToAdd = "[CODE]";
document.form.code.value = " ����*";
Code = 1;
} else {
ToAdd = "[/CODE]";
document.form.code.value = " ���� ";
Code = 0;
}
commentWrite(ToAdd);
}
function Cquote() {
fontbegin="[QUOTE]";
fontend="[/QUOTE]";
fontchuli();
}
function Cbold() {
fontbegin="[B]";
fontend="[/B]";
fontchuli();
}
function Citalic() {
fontbegin="[I]";
fontend="[/I]";
fontchuli();
}
function Cunder() {
fontbegin="[U]";
fontend="[/U]";
fontchuli();
}
function Ccenter() {
fontbegin="[align=center]";
fontend="[/align]";
fontchuli();
}
function Cstrike() {
fontbegin="[strike]";
fontend="[/strike]";
fontchuli();
}
function Csound() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_sound, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("����"+FoundErrors);
return;
}
var ToAdd = "[SOUND]"+enterURL+"[/SOUND]";
document.oblogform.ubbedit.value+=ToAdd;
document.oblogform.ubbedit.focus();
}

helpstat = false;
stprompt = true;
basic = false;
function thelp(swtch){
	if (swtch == 1){
		basic = false;
		stprompt = false;
		helpstat = true;
	} else if (swtch == 0) {
		helpstat = false;
		stprompt = false;
		basic = true;
	} else if (swtch == 2) {
		helpstat = false;
		basic = false;
		stprompt = true;
	}
}

function Cswf() {
 	if (helpstat){
		alert("Flash\nFlash ����.\n�÷�: [flash=���, �߶�]Flash �ļ��ĵ�ַ[/flash]");
	} else if (basic) {
		AddTxt="[flash=500,350][/flash]";
		AddText(AddTxt);
	} else {                  
		txt2=prompt("flash��ȣ��߶�","500,350"); 
		if (txt2!=null) {
                txt=prompt("Flash �ļ��ĵ�ַ","http://");
		if (txt!=null) {
                          if (txt2=="") {             
			AddTxt="[flash=500,350]"+txt;
			AddText(AddTxt);
			AddTxt="[/flash]";
			AddText(AddTxt);
               } else {
		        AddTxt="[flash="+txt2+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/flash]";
			AddText(AddTxt);
		 }        
	    }  
       }
    }
}

function Crm() {
	if (helpstat) {
               alert("realplay\n����realplay�ļ�.\n�÷�: [rm=���, �߶�]�ļ���ַ[/rm]");
	} else if (basic) {
		AddTxt="[rm=500,350][/rm]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("��Ƶ�Ŀ�ȣ��߶�","500,350"); 
		if (txt2!=null) {
			txt=prompt("��Ƶ�ļ��ĵ�ַ","������");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[rm=500,350]"+txt;
					AddText(AddTxt);
					AddTxt="[/rm]";
					AddText(AddTxt);
				} else {
					AddTxt="[rm="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/rm]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function Cwmv() {
	if (helpstat) {
               alert("Media Player\n����Media Player�ļ�.\n�÷�: [mp=���, �߶�]�ļ���ַ[/mp]");
	} else if (basic) {
		AddTxt="[mp=500,350][/mp]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("��Ƶ�Ŀ�ȣ��߶�","500,350"); 
		if (txt2!=null) {
			txt=prompt("��Ƶ�ļ��ĵ�ַ","������");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[mp=500,350]"+txt;
					AddText(AddTxt);
					AddTxt="[/mp]";
					AddText(AddTxt);
				} else {
					AddTxt="[mp="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/mp]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function Cmov() {
	if (helpstat) {
               alert("QuickTime\n����QuickTime�ļ�.\n�÷�: [qt=���, �߶�]�ļ���ַ[/qt]");
	} else if (basic) {
		AddTxt="[qt=500,350][/qt]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("��Ƶ�Ŀ�ȣ��߶�","500,350"); 
		if (txt2!=null) {
			txt=prompt("��Ƶ�ļ��ĵ�ַ","������");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[qt=500,350]"+txt;
					AddText(AddTxt);
					AddTxt="[/qt]";
					AddText(AddTxt);
				} else {
					AddTxt="[qt="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/qt]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function Cdir() {
	if (helpstat) {
               alert("Shockwave\n����Shockwave�ļ�.\n�÷�: [dir=���, �߶�]�ļ���ַ[/dir]");
	} else if (basic) {
		AddTxt="[dir=500,350][/dir]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("Shockwave�ļ��Ŀ�ȣ��߶�","500,350"); 
		if (txt2!=null) {
			txt=prompt("Shockwave�ļ��ĵ�ַ","�������ַ");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[dir=500,350]"+txt;
					AddText(AddTxt);
					AddTxt="[/dir]";
					AddText(AddTxt);
				} else {
					AddTxt="[dir="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/dir]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}



function Cra() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_ra, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("����"+FoundErrors);
return;
}
var ToAdd = "[RA]"+enterURL+"[/RA]";
document.oblogform.ubbedit.value+=ToAdd;
document.oblogform.ubbedit.focus();
}



function Cfanzi() {
fontbegin="[xray]";
fontend="[/xray]";
fontchuli();
}

function Cwma() {
var FoundErrors = '';
var enterURL   = prompt(text_enter_wma, "http://");
if (!enterURL) {
FoundErrors += "\n" + error_no_url;
}
if (FoundErrors) {
alert("����"+FoundErrors);
return;
}
var ToAdd = "[wma]"+enterURL+"[/wma]";
document.oblogform.ubbedit.value+=ToAdd;
document.oblogform.ubbedit.focus();
}


function Cmarquee() {
fontbegin="[move]";
fontend="[/move]";
fontchuli();
}
function Cfly() {
fontbegin="[fly]";
fontend="[/fly]";
fontchuli();
}

function paste(text) {
	if (opener.document.oblogform.ubbedit.createTextRange && opener.document.oblogform.ubbedit.caretPos) {      
		var caretPos = opener.document.oblogform.ubbedit.caretPos;      
		caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
		text + ' ' : text;
	}
	else opener.document.oblogform.ubbedit.value += text;
	opener.document.oblogform.ubbedit.focus(caretPos);
}

function showsize(size){
fontbegin="[size="+size+"]";
fontend="[/size]";
fontchuli();
}

function showfont(font){
fontbegin="[face="+font+"]";
fontend="[/face]";
fontchuli();
}

function showcolor(color){
fontbegin="[color="+color+"]";
fontend="[/color]";
fontchuli();
}

function fontchuli(){
if ((document.selection)&&(document.selection.type == "Text")) {
var range = document.selection.createRange();
var ch_text=range.text;
range.text = fontbegin + ch_text + fontend;
} 
else {
document.oblogform.ubbedit.value=fontbegin+document.oblogform.ubbedit.value+fontend;
document.oblogform.ubbedit.focus();
}
}

function Cguang() {
var FoundErrors = '';
var enterSET   = prompt(text_enter_guang1, "255,red,2");
var enterTxT   = prompt(text_enter_guang2, "����");
if (!enterSET)    {
FoundErrors += "\n" + error_no_gset;
}
if (!enterTxT)    {
FoundErrors += "\n" + error_no_gtxt;
}
if (FoundErrors)  {
alert("����"+FoundErrors);
return;
}
var ToAdd = "[glow="+enterSET+"]"+enterTxT+"[/glow]";
document.oblogform.ubbedit.value+=ToAdd;
document.oblogform.ubbedit.focus();
}

function Cying() {
var FoundErrors = '';
var enterSET   = prompt(text_enter_guang1, "255,blue,1");
var enterTxT   = prompt(text_enter_guang2, "����");
if (!enterSET)    {
FoundErrors += "\n" + error_no_gset;
}
if (!enterTxT)    {
FoundErrors += "\n" + error_no_gtxt;
}
if (FoundErrors)  {
alert("����"+FoundErrors);
return;
}
var ToAdd = "[SHADOW="+enterSET+"]"+enterTxT+"[/SHADOW]";
document.oblogform.ubbedit.value+=ToAdd;
document.oblogform.ubbedit.focus();
}

ie = (document.all)? true:false
if (ie){
function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.oblogform.submit();}}
}

function popreview()
{
document.preview.title.value=document.oblogform.title.value;
document.preview.content.value=document.oblogform.ubbedit.value;
var popupWin = window.open('user_post.asp?action=preview', 'preview_page', 'scrollbars=yes,width=550,height=450');
document.preview.submit()
}

//