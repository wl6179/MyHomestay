<link rel="STYLESHEET" type="text/css" href="images/edit.css">
<Script Src="images/DhtmlEdit.js"></Script>
<table id="oblog_Container" class="oblog_Body" height=100% width=350 cellpadding=1 cellspacing=0 border=0 >
  <tr> 
    <td  height="10"> <table cellpadding=0 cellspacing=0 >
        <tr class="yToolbar" ID="ExtToolbar0" > 
          <td> <select language="javascript" class="oblog_TBGen" id="FontSize" onchange="FormatText('fontsize',this[this.selectedIndex].value);">
              <option class="heading" selected>字号 
              <option value="1">1 
              <option value="2">2 
              <option value="3">3 
              <option value="4">4 
              <option value="5">5 
              <option value="6">6 
              <option value="7">7</option>
            </select> 
          <td class="oblog_Btn" TITLE="加粗" LANGUAGE="javascript" onclick="FormatText('bold', '');ondrag='return false;'" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn'; > 
            <img class="oblog_Ico" src="images/bold.gif" WIDTH="16" HEIGHT="16" unselectable="on"> </td>
          <td class="oblog_Btn" TITLE="斜体" LANGUAGE="javascript" onclick="FormatText('italic', '');ondrag='return false;'" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn'; > 
            <img class="oblog_Ico" src="images/italic.gif" WIDTH="16" HEIGHT="16" unselectable="on"> </td>
          <td class="oblog_Btn" TITLE="下划线" LANGUAGE="javascript" onclick="FormatText('underline', '');ondrag='return false;'" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn'; > 
            <img class="oblog_Ico" src="images/underline.gif" WIDTH="16" HEIGHT="16" unselectable="on"> </td>
          <td class="oblog_Btn" TITLE="取消格式" LANGUAGE="javascript" onclick="FormatText('RemoveFormat', '');ondrag='return false;'" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn'; > 
            <img class="oblog_Ico" src="images/removeformat.gif" WIDTH="16" HEIGHT="16" unselectable="on"> </td>
          <td class="oblog_Btn" TITLE="左对齐" NAME="Justify" LANGUAGE="javascript" onclick="FormatText('justifyleft', '');ondrag='return false;'" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn'; > 
            <img class="oblog_Ico" src="images/aleft.gif" WIDTH="16" HEIGHT="16" unselectable="on"> </td>
          <td class="oblog_Btn" TITLE="居中" NAME="Justify" LANGUAGE="javascript" onclick="FormatText('justifycenter', '');ondrag='return false;'" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn'; > 
            <img class="oblog_Ico" src="images/center.gif" WIDTH="16" HEIGHT="16" unselectable="on"> </td>
          <td class="oblog_Btn" TITLE="右对齐" NAME="Justify" LANGUAGE="javascript" onclick="FormatText('justifyright', '');ondrag='return false;'" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn'; > 
            <img class="oblog_Ico" src="images/aright.gif" WIDTH="16" HEIGHT="16" unselectable="on"> </td>
          <td class="oblog_Btn" TITLE="插入表情" LANGUAGE="javascript" onclick="oblog_foremot()" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn'; > 
            <img class="oblog_Ico" src="images/smiley.gif" WIDTH="16" HEIGHT="16" unselectable="on"> </td>
          <td id="forecolor" name=forecolor class="oblog_Btn" TITLE="字体颜色" LANGUAGE="javascript" onclick="oblog_foreColor();" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn'; > 
            <img class="oblog_Ico" src="images/fgcolor.gif" WIDTH="16" HEIGHT="16" unselectable="on" > </td>
          <td id="backcolor" class="oblog_Btn" TITLE="字体背景颜色" LANGUAGE="javascript" onclick="oblog_backColor();ondrag='return false;'" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn';> 
            <img class="oblog_Ico" src="images/fbcolor.gif" WIDTH="16" HEIGHT="16" unselectable="on"> </td>
          <td class="oblog_Btn" TITLE="插入超级链接" LANGUAGE="javascript" onclick="oblog_forlink();ondrag='return false;'" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn'; > 
            <img class="oblog_Ico" src="images/wlink.gif" WIDTH="18" HEIGHT="18" unselectable="on"> </td>
          <td class="oblog_Btn" TITLE="去掉超级链接" LANGUAGE="javascript" onclick="FormatText('Unlink');ondrag='return false;'" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn'; > 
            <img class="oblog_Ico" src="images/unlink.gif" WIDTH="16" HEIGHT="16" unselectable="on"> </td>
          <td class="oblog_Btn" TITLE="清理代码" LANGUAGE="javascript" onclick="oblog_CleanCode();ondrag='return false;'" onmouseover=this.className='oblog_BtnMouseOverUp'; onmouseout=this.className='oblog_Btn';> 
            <img class="oblog_Ico" src="images/cleancode.gif" WIDTH="16" HEIGHT="16"></td>
        </tr>
      </table></tr>
  <tr> 
    <td height="100%" id=PostiFrame> <iframe class="oblog_Composition" ID="oblog_Composition" MARGINHEIGHT="5" MARGINWIDTH="5" width="100%" height="100%"></iframe> </td>
     </tr>
  <tr >
    <td height="20"><TABLE CELLPADDING=0 CELLSPACING=0 BORDER=0 width='100%'>
        <TR> 
          <TD><div align="right"><a href="javascript:oblog_Size(-100)"><img src="images/minus.gif" unselectable="on" border='0' height="20"></a> <a href="javascript:oblog_Size(100)"><img src="images/plus.gif" unselectable="on" border='0' height="20"></a></div></TD>
          <TD width='20'></TD>
        </TR>
      </TABLE></td>
  </tr>
</table>
<Script>
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
function initx(){
//IframeID.document.body.innerHTML=quote
}
initx();
</Script>
<Script Src="images/editor.js"></Script>