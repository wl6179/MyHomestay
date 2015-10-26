<div id=demo style=overflow:hidden;height:500;width:300>

<div id=demo1>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/C_02.jpg" width="120" height="93" /></td>
                <td><img src="images/C_03.jpg" width="120" height="93" /></td>
              </tr>
            </table>
      </td>
      </tr>
      <tr>
        <td background="images/dot_line.gif"><img src="images/blank.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td><img src="images/C_04.jpg" width="120" height="93" /></td>
              <td><img src="images/C_05.jpg" width="120" height="93" /></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td background="images/dot_line.gif"><img src="images/blank.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/C_06.jpg" width="120" height="93" /></td>
                <td><img src="images/C_07.jpg" width="120" height="93" /></td>
              </tr>
            </table>
  </td>
      </tr>
      <tr>
        <td background="images/dot_line.gif"><img src="images/blank.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/C_08.jpg" width="120" height="93" /></td>
                <td><img src="images/C_09.jpg" width="120" height="93" /></td>
              </tr>
            </table>
      </td>
      </tr>
      <tr>
        <td background="images/dot_line.gif"><img src="images/blank.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/C_10.jpg" width="120" height="93" /></td>
                <td><img src="images/C_11.jpg" width="120" height="93" /></td>
              </tr>
            </table>
       </td>
      </tr>
      <tr>
        <td background="images/dot_line.gif"><img src="images/blank.gif" width="1" height="1"></td>
      </tr>
    </table>
      <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><img src="images/C_12.jpg" width="120" height="93" /></td>
                  <td><img src="images/C_13.jpg" width="120" height="93" /></td>
                </tr>
              </table>
       </td>
        </tr>
        <tr>
          <td background="images/dot_line.gif"><img src="images/blank.gif" width="1" height="1"></td>
        </tr>
        <tr>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/C_14.jpg" width="120" height="93" /></td>
                <td><img src="images/C_15.jpg" width="120" height="93" /></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td background="images/dot_line.gif"><img src="images/blank.gif" width="1" height="1"></td>
        </tr>
        <tr>
          <td>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><img src="images/C_16.jpg" width="120" height="93" /></td>
                  <td><img src="images/C_17.jpg" width="120" height="93" /></td>
                </tr>
              </table>
         </td>
        </tr>
        <tr>
          <td background="images/dot_line.gif"><img src="images/blank.gif" width="1" height="1"></td>
        </tr>
        <tr>
          <td>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><img src="images/C_19.jpg" width="120" height="93" /></td>
                  <td><img src="images/C_20.jpg" width="120" height="93" /></td>
                </tr>
              </table>
        </td>
        </tr>
        <tr>
          <td background="images/dot_line.gif"><img src="images/blank.gif" width="1" height="1"></td>
        </tr>
        <tr>
          <td>
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><img src="images/C_21.jpg" width="120" height="93" /></td>
                  <td><img src="images/C_22.jpg" width="120" height="93" /></td>
                </tr>
              </table>
         </td>
        </tr>
        <tr>
          <td background="images/dot_line.gif"><img src="images/blank.gif" width="1" height="1"></td>
        </tr>
      </table>
</td>
  </tr>
</table>
</div>

<div id=demo2></div>

</div>


<div align="left">
	<script>
		var speed=30
		demo2.innerHTML=demo1.innerHTML
		function Marquee(){
		if(demo2.offsetTop-demo.scrollTop<=0)
		demo.scrollTop-=demo1.offsetHeight
		else{
		demo.scrollTop++
		}
		}
		var MyMar=setInterval(Marquee,speed)
		demo.onmouseover=function() {clearInterval(MyMar)}
		demo.onmouseout=function() {MyMar=setInterval(Marquee,speed)}
	</script>
</div>
