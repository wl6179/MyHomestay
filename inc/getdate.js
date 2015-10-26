//Begin     wl    目的:调用date.html代码的代码.
function getDate(strdate){
var Arguments = new Array(1);
Arguments[0]=strdate;
var strResult=showModalDialog("inc/date.html",Arguments,"center:yes;resizable:no;dialogHeight:235px;dialogWidth:230px;help:no;status:no"); 
if (strResult!=null){
  return strResult;
}
else{
  return "";
}
}
//End      wl