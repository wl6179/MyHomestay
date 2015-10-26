var IsIE5=document.all;
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
		if(e.type == 'checkbox' && e.name != 'chkAll'){
			var objParentDiv = e.parentNode.parentNode;
			e.checked ? fSetBg(objParentDiv) : fReBg(objParentDiv);
		}
    }
}

function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}

function submitdate()
{
	var date=document.oblogform.selecty1.value+"-"+document.oblogform.selectm1.value+"-"+document.oblogform.selectd1.value
	var datereg=/^(\d{4})-(\d{1,2})-(\d{1,2})$/
	var datareg=/^(\d){1,2}$/
	if (!datereg.test(date)){
	  alert("开始时间输入格式错误");
	   return false;
	 }
	var r=date.match(datereg)
	var d=new Date(r[1],r[2]-1,r[3])
	if (!(d.getFullYear()==r[1]&&d.getMonth()==r[2]-1&&d.getDate()==r[3])){
	  alert("开始时间输入格式错误");
	   return false;
	 }
	 
	 var edate=document.oblogform.selecty2.value+"-"+document.oblogform.selectm2.value+"-"+document.oblogform.selectd2.value
	var datereg=/^(\d{4})-(\d{1,2})-(\d{1,2})$/
	var datareg=/^(\d){1,2}$/
	if (!datereg.test(edate)){
	  alert("结束时间输入格式错误");
	   return false;
	 }
	var er=edate.match(datereg)
	var ed=new Date(er[1],er[2]-1,er[3])
	if (!(ed.getFullYear()==er[1]&&ed.getMonth()==er[2]-1&&ed.getDate()==er[3])){
	  alert("结束时间输入格式错误");
	   return false;
	 }	
	
	date = date.replace(/\-/g,"\/");
	edate = edate.replace(/\-/g,"\/");	
	if ((new Date(date) > new Date(edate))){
		alert("结束日期不能小于开始日期!");
	return false;
	}
	 
	return true;
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

function chkcopy()
{
	if (IsIE5){
		var IframeID=frames["oblog_Composition"];
	}
	else{
		var IframeID=document.getElementById("oblog_Composition").contentWindow;
	}
	if(IframeID !=null){ 
		//document.oblogform.edit.value=IframeID.document.body.innerHTML;
		//var tmptext=document.oblogform.edit.createTextRange(); 
		//tmptext.execCommand("Copy");
		var tmptext=IframeID.document.body.innerHTML;
		if (tmptext!=null) {
		window.clipboardData.setData("Text",tmptext);
		}
	}
	
	var ubbid=document.getElementById("ubbedit");
	if(ubbid != null){ 
		if (document.oblogform.ubbedit.value!=null) {
		var tmptext1=document.oblogform.ubbedit.createTextRange(); 
		tmptext1.execCommand("Copy");
		}
	}

}

function fSetBg(obj){
	//obj.style.backgroundColor = '#cccccc';
	if (IsIE5){
	obj.className='list_content_mouserover';
	}
}
function fReBg(obj){
	if (IsIE5){
		var objChildCheck = document.getElementById("list")? obj.children[0].children[0] : obj.childNodes[1].childNodes[1];
		if(objChildCheck.checked){
			return false;
		}
		//obj.style.backgroundColor = '';
		obj.className='list_content';
	}
}

//Nav_Function
function menuFix() {
    var sfEls = document.getElementById("nav").getElementsByTagName("li");
    for (var i=0; i<sfEls.length; i++) {
        sfEls[i].onmouseover=function() {
        this.className+=(this.className.length>0? " ": "") + "sfhover";
        }
        sfEls[i].onMouseDown=function() {
        this.className+=(this.className.length>0? " ": "") + "sfhover";
        }
        sfEls[i].onMouseUp=function() {
        this.className+=(this.className.length>0? " ": "") + "sfhover";
        }
        sfEls[i].onmouseout=function() {
        this.className=this.className.replace(new RegExp("( ?|^)sfhover\\b"),

"");
        }
    }
}
//window.onload=menuFix;
//End of Nav_Function

